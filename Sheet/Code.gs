const DM = {
  // Вставьте ID четырёх Google-таблиц. ID находится между /d/ и /edit в URL.
  centralFileId: '',
  plusFileIds: {
    'Артём': '',
    'Рами': '',
    'Немат': ''
  },
  practitioners: ['Артём', 'Рами', 'Немат'],
  common: 'Общий',
  cw: 'CW',
  ranking: 'Рейтинг',
  logs: 'Логи',
  plusLogs: 'Логи',
  practiceCount: 15,
  studentsPerGroup: 30,
  plusHeaderRow: 3,
  plusFirstRow: 4,
  plusCoefficientCol: 4,
  plusTotalCol: 5,
  taskFirstCol: 6,
  taskLastCol: 35,
  commonFirstRow: 4,
  commonLastRow: 93,
  logFirstRow: 4,
  logMaxRow: 3006,
  timeZone: 'Europe/Moscow'
};

function plusLayout_(practitioner) {
  if (practitioner === 'Рами') {
    return {
      presenceCol: 0,
      coefficientCol: 3,
      totalCol: 4,
      taskFirstCol: 5,
      taskLastCol: 34
    };
  }
  return {
    presenceCol: 2,
    coefficientCol: DM.plusCoefficientCol,
    totalCol: DM.plusTotalCol,
    taskFirstCol: DM.taskFirstCol,
    taskLastCol: DM.taskLastCol
  };
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('DM система')
    .addItem('Настроить 4 файла', 'installDmSystem')
    .addItem('Синхронизировать группы', 'syncGroups_')
    .addItem('Синхронизировать таблицу', 'syncTable')
    .addToUi();
}

function installDmSystem() {
  validateConfig_();
  setupTriggers_();
  const central = SpreadsheetApp.openById(getCentralId_());
  const plusFiles = openPlusFiles_();
  setSystemTimeZones_(central, plusFiles);
  syncGroups_(central, plusFiles);
  SpreadsheetApp.getActive().toast('Центральная таблица и три файла практиков настроены.');
}

// Установочный триггер должен быть создан для всех четырёх таблиц.
function setupTriggers_() {
  const ids = [getCentralId_()].concat(DM.practitioners.map(name => DM.plusFileIds[name]));
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'handleEdit_') ScriptApp.deleteTrigger(trigger);
  });
  ids.filter(Boolean).forEach(id => {
    ScriptApp.newTrigger('handleEdit_').forSpreadsheet(id).onEdit().create();
  });
}

// Это единственный обработчик изменений во всех четырёх Google-таблицах.
function handleEdit_(event) {
  if (!event || !event.range || !event.source) return;
  const sourceId = event.source.getId();
  if (sourceId === getCentralId_()) {
    handleCentralEdit_(event);
    return;
  }
  const practitioner = DM.practitioners.find(name => DM.plusFileIds[name] === sourceId);
  if (practitioner) handlePlusEdit_(event, practitioner);
}

function handleCentralEdit_(event) {
  const sheet = event.range.getSheet();
  const row = event.range.getRow();
  const col = event.range.getColumn();
  const central = event.source;

  if (sheet.getName() === DM.common && row >= DM.commonFirstRow && row <= DM.commonLastRow) {
    if (col === 1) {
      const oldName = String(event.oldValue || '').trim();
      const newName = String(sheet.getRange(row, 1).getDisplayValue()).trim();
      if (oldName && !newName) {
        sheet.getRange(row, 1).setValue(oldName);
        central.toast('ФИО нельзя очищать. Укажите новое имя студента.');
        return;
      }
      if (oldName && newName && oldName !== newName) {
        const practitioner = String(sheet.getRange(row, 2).getDisplayValue()).trim();
        const result = renameStudentEverywhere_(central, oldName, newName, practitioner, row);
        if (!result.ok) {
          sheet.getRange(row, 1).setValue(oldName);
          central.toast(result.message);
          return;
        }
      }
      syncGroups_();
      return;
    }
    if (col === 2) appendTransferLog_(central, row, event);
    if (col === 4) appendDepartureLog_(central, row);
    if (col >= 1 && col <= 4) {
      syncGroups_();
      return;
    }
    if (col === 10 || col === 11 || col === 12 || col === 13) {
      refreshCentral_();
      return;
    }
  }

  if (sheet.getName() === DM.cw ||
      (sheet.getName() === DM.logs && row >= DM.logFirstRow && (col === 8 || col === 10))) {
    refreshCentral_();
  }
}

function handlePlusEdit_(event, practitioner) {
  const sheet = event.range.getSheet();
  const cell = event.range;
  const row = cell.getRow();
  const col = cell.getColumn();
  const layout = plusLayout_(practitioner);
  const plusFirst = DM.plusFirstRow;
  const plusLast = plusFirst + DM.studentsPerGroup - 1;

  if (col === 1 && row >= plusFirst && row <= plusLast) {
    const oldName = String(event.oldValue || '').trim();
    const newName = String(cell.getDisplayValue()).trim();
    if (oldName && !newName) {
      cell.setValue(oldName);
      event.source.toast('ФИО нельзя очищать. Укажите новое имя студента.');
      return;
    }
    if (!oldName || !newName || oldName === newName) return;
    const central = SpreadsheetApp.openById(getCentralId_());
    const result = renameStudentEverywhere_(central, oldName, newName, practitioner);
    if (!result.ok) {
      cell.setValue(oldName);
      event.source.toast(result.message);
      return;
    }
    const knownFiles = {};
    knownFiles[practitioner] = event.source;
    syncGroups_(central, knownFiles);
    return;
  }
  if (col < layout.taskFirstCol || col > layout.taskLastCol) return;
  cell.clearNote();

  const practiceNo = practiceFromSheet_(sheet);
  if (!practiceNo) return;
  const plusHeader = DM.plusHeaderRow;
  if (row < plusFirst || row > plusLast) return;

  const value = String(cell.getDisplayValue()).trim();
  const oldValue = event.range.getNumRows() === 1 && event.range.getNumColumns() === 1
    ? String(event.oldValue || '').trim() : '';
  const studentName = String(sheet.getRange(row, 1).getDisplayValue()).trim();
  const taskName = String(sheet.getRange(plusHeader, col).getDisplayValue()).trim();
  if (!studentName || !taskName) return;

  const normalizedOldValue = oldPlusValue_(oldValue);
  if (value === '+' && normalizedOldValue !== '+') {
    appendPlusHistory_(event.source, studentName, practiceNo, taskName, 'Поставил');
  } else if (normalizedOldValue === '+' && value !== '+') {
    appendPlusHistory_(event.source, studentName, practiceNo, taskName, 'Убрал');
  }

  const central = SpreadsheetApp.openById(getCentralId_());
  const knownFiles = {};
  knownFiles[practitioner] = event.source;
  const practitionerNo = DM.practitioners.indexOf(practitioner) + 1;

  if (value === '1') {
    removeLogRecord_(central, practitioner, practiceNo, studentName, taskName, true);
    if (oldValue !== '1') {
      const result = appendNormalLog_(central, event.source, sheet, practitionerNo,
        practitioner, practiceNo, studentName, taskName);
      if (!result.ok) {
        cell.setValue(oldValue === '+' ? '+' : '');
        event.source.toast(result.message);
        return;
      }
    }
    refreshCentral_(central, knownFiles);
    return;
  }

  if (value === '-') {
    removeLogRecord_(central, practitioner, practiceNo, studentName, taskName, false);
    if (oldValue !== '-') appendPenaltyLog_(central, practitioner, practiceNo, studentName, taskName);
    refreshCentral_(central, knownFiles);
    return;
  }

  if (value === '' || value === '+') {
    if (oldValue === '1' || oldValue === '-') {
      removeLogRecord_(central, practitioner, practiceNo, studentName, taskName, false);
    }
    refreshCentral_(central, knownFiles);
  }
}

function renameStudentEverywhere_(central, oldName, newName, practitioner, knownRow) {
  const common = central.getSheetByName(DM.common);
  const commonRows = common.getRange(DM.commonFirstRow, 1, DM.commonLastRow - DM.commonFirstRow + 1, 2).getDisplayValues();
  let sourceIndex = -1;
  commonRows.forEach((row, index) => {
    if (sourceIndex < 0 && String(row[0]).trim() === oldName &&
        (!practitioner || String(row[1]).trim() === practitioner)) sourceIndex = index;
  });
  if (sourceIndex < 0 && knownRow >= DM.commonFirstRow && knownRow <= DM.commonLastRow) {
    sourceIndex = knownRow - DM.commonFirstRow;
  }
  if (sourceIndex < 0) {
    return { ok: false, message: 'Старое ФИО не найдено в центральной таблице.' };
  }
  const duplicate = commonRows.some((row, index) => index !== sourceIndex && String(row[0]).trim() === newName);
  if (duplicate) {
    return { ok: false, message: 'Студент с таким ФИО уже есть в центральной таблице.' };
  }

  common.getRange(DM.commonFirstRow + sourceIndex, 1).setValue(newName);

  const logLastRow = findLastLogRow_(central.getSheetByName(DM.logs));
  if (logLastRow >= DM.logFirstRow) {
    const logs = central.getSheetByName(DM.logs);
    const values = logs.getRange(DM.logFirstRow, 1, logLastRow - DM.logFirstRow + 1, 11).getValues();
    let changed = false;
    values.forEach(row => {
      if (String(row[0]).trim() === oldName) {
        row[0] = newName;
        changed = true;
      }
    });
    if (changed) logs.getRange(DM.logFirstRow, 1, values.length, 11).setValues(values);
  }

  DM.practitioners.forEach(name => {
    const fileId = DM.plusFileIds[name];
    if (!fileId) return;
    const file = SpreadsheetApp.openById(fileId);
    const localLogs = file.getSheetByName(DM.plusLogs);
    if (!localLogs) return;
    const lastRow = localLogs.getLastRow();
    if (lastRow < 2) return;
    const values = localLogs.getRange(2, 1, lastRow - 1, 5).getValues();
    let changed = false;
    values.forEach(row => {
      if (String(row[0]).trim() === oldName) {
        row[0] = newName;
        changed = true;
      }
    });
    if (changed) localLogs.getRange(2, 1, values.length, 5).setValues(values);
  });

  return { ok: true };
}

function appendNormalLog_(central, plusFile, plusSheet, practitionerNo,
                         practitioner, practiceNo, studentName, taskName) {
  if (!isStudentActive_(central, practitioner, studentName)) {
    return { ok: false, message: 'Студент неактивен или уже находится у другого практика.' };
  }

  const logs = central.getSheetByName(DM.logs);
  const lastRow = findLastLogRow_(logs);
  const rows = lastRow >= DM.logFirstRow
    ? logs.getRange(DM.logFirstRow, 1, lastRow - DM.logFirstRow + 1, 9).getDisplayValues() : [];
  let studentExit = 0;
  let pairCount = 0;
  let duplicateCount = 0;

  rows.forEach(row => {
    const penalty = String(row[5]).trim() === '' && String(row[6]).trim() === '' && Number(row[7]) === -2;
    const departureOrTransfer = ['ушел на базу', 'переведен из'].some(status => String(row[8]).indexOf(status) === 0);
    if (String(row[0]).trim() !== studentName || penalty || departureOrTransfer) return;
    studentExit++;
    if (Number(row[2]) === practiceNo) pairCount++;
    if (Number(row[2]) === practiceNo && tasksEqual_(row[3], taskName)) duplicateCount++;
  });

  if (pairCount >= 2) return { ok: false, message: 'У студента уже два выхода за эту практику.' };
  if (duplicateCount > 0) return { ok: false, message: 'Эта задача уже есть в логах.' };

  const plusFirst = DM.plusFirstRow;
  const studentRow = findStudentRow_(plusSheet, studentName, plusFirst, plusFirst + DM.studentsPerGroup - 1);
  if (!studentRow) return { ok: false, message: 'Студент не найден в файле практика.' };
  const layout = plusLayout_(practitioner);
  const coefficient = Number(plusSheet.getRange(studentRow, layout.coefficientCol).getValue()) || 0.5;
  const exitNo = studentExit + 1;
  const base = ladderPoints_(exitNo);
  const newRow = Math.max(lastRow + 1, DM.logFirstRow);
  logs.getRange(newRow, 1, 1, 11).setValues([[
    studentName, practitioner, practiceNo, taskName, exitNo,
    base, coefficient, '', exitNo > 15 ? 'Дополнительный выход' : 'Назначен', '', new Date()
  ]]);
  logs.getRange(newRow, 8).setFormula(
    `=IF(OR(F${newRow}="",G${newRow}=""),"",F${newRow}*G${newRow}+IF(J${newRow}="",0,J${newRow}))`
  );
  logs.getRange(newRow, 11).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  return { ok: true, message: `Начислено D: ${base * coefficient}` };
}

function appendPlusHistory_(plusFile, studentName, practiceNo, taskName, state) {
  const logs = plusFile.getSheetByName(DM.plusLogs);
  if (!logs) return;
  const lastRow = Math.max(logs.getLastRow() + 1, 2);
  logs.getRange(lastRow, 1, 1, 5).setValues([[
    studentName, practiceNo, taskName, state, new Date()
  ]]);
  logs.getRange(lastRow, 5).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

function oldPlusValue_(value) {
  return String(value == null ? '' : value).trim().replace(/^'/, '');
}

function appendPenaltyLog_(central, practitioner, practiceNo, studentName, taskName) {
  const logs = central.getSheetByName(DM.logs);
  const newRow = Math.max(findLastLogRow_(logs) + 1, DM.logFirstRow);
  logs.getRange(newRow, 1, 1, 11).setValues([[
    studentName, practitioner, practiceNo, taskName, '', '', '', -2, '', '', new Date()
  ]]);
  logs.getRange(newRow, 11).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

function removeLogRecord_(central, practitioner, practiceNo, studentName, taskName, onlyPenalty) {
  const logs = central.getSheetByName(DM.logs);
  const lastRow = findLastLogRow_(logs);
  if (lastRow < DM.logFirstRow) return;
  const rows = logs.getRange(DM.logFirstRow, 1, lastRow - DM.logFirstRow + 1, 9).getDisplayValues();
  const ranges = [];
  rows.forEach((row, index) => {
    const penalty = String(row[5]).trim() === '' && String(row[6]).trim() === '' && Number(row[7]) === -2;
    const same = String(row[0]).trim() === studentName &&
      Number(row[2]) === practiceNo &&
      tasksEqual_(row[3], taskName) && (!onlyPenalty || penalty);
    if (same) {
      const sheetRow = DM.logFirstRow + index;
      ranges.push(`A${sheetRow}:K${sheetRow}`);
    }
  });
  if (ranges.length) logs.getRangeList(ranges).clearContent();
}

function appendTransferLog_(central, commonRow, event) {
  if (!event || !event.oldValue || event.range.getNumRows() !== 1 || event.range.getNumColumns() !== 1) return;
  const common = central.getSheetByName(DM.common);
  const student = String(common.getRange(commonRow, 1).getDisplayValue()).trim();
  const next = String(common.getRange(commonRow, 2).getDisplayValue()).trim();
  const previous = String(event.oldValue).trim();
  if (!student || !next || !previous || previous === next) return;
  appendStatusLog_(central, student, next, `переведен из ${previous} в ${next}`);
}

function appendDepartureLog_(central, commonRow) {
  const common = central.getSheetByName(DM.common);
  if (String(common.getRange(commonRow, 4).getDisplayValue()).trim() !== 'Нет') return;
  const student = String(common.getRange(commonRow, 1).getDisplayValue()).trim();
  const practitioner = String(common.getRange(commonRow, 2).getDisplayValue()).trim();
  if (!student || !practitioner || hasStatusLog_(central, student, practitioner, 'ушел на базу')) return;
  appendStatusLog_(central, student, practitioner, 'ушел на базу');
}

function appendStatusLog_(central, student, practitioner, status) {
  const logs = central.getSheetByName(DM.logs);
  const newRow = Math.max(findLastLogRow_(logs) + 1, DM.logFirstRow);
  logs.getRange(newRow, 1, 1, 11).setValues([[
    student, practitioner, '', '', '', '', '', '', status, '', new Date()
  ]]);
  logs.getRange(newRow, 11).setNumberFormat('yyyy-mm-dd hh:mm:ss');
}

function hasStatusLog_(central, student, practitioner, status) {
  const logs = central.getSheetByName(DM.logs);
  const lastRow = findLastLogRow_(logs);
  if (lastRow < DM.logFirstRow) return false;
  return logs.getRange(DM.logFirstRow, 1, lastRow - DM.logFirstRow + 1, 9).getDisplayValues()
    .some(row => String(row[0]).trim() === student && String(row[1]).trim() === practitioner && String(row[8]).trim() === status);
}

function syncTable() {
  validateConfig_();
  syncGroups_();
  SpreadsheetApp.getActive().toast('Центральная таблица синхронизирована.');
}

function refreshCentral_legacy() {
  const central = SpreadsheetApp.openById(getCentralId_());
  SpreadsheetApp.flush();
  refreshLogCoefficients_(central);
  SpreadsheetApp.flush();
  const common = central.getSheetByName(DM.common);
  for (let row = DM.commonFirstRow; row <= DM.commonLastRow; row++) {
    common.getRange(row, 5).setFormula(
      `=IF(A${row}="","",SUMIF('${DM.logs}'!$A$${DM.logFirstRow}:$A$${DM.logMaxRow},A${row},'${DM.logs}'!$H$${DM.logFirstRow}:$H$${DM.logMaxRow}))`
    );
    common.getRange(row, 6).setFormula(`=IF(A${row}="","",SUMIF('${DM.cw}'!$A:$A,A${row},'${DM.cw}'!$D:$D))`);
    common.getRange(row, 7).setFormula(
      `=IF(A${row}="","",IF(AND(E${row}+IF(I${row}="",0,I${row})>=55,F${row}+IF(H${row}="",0,H${row})>=24),"S",IF(AND(E${row}+IF(I${row}="",0,I${row})>=35,F${row}+IF(H${row}="",0,H${row})>=14),"A",IF(E${row}+IF(I${row}="",0,I${row})>=17,"B","F"))))`
    );
    common.getRange(row, 11).setFormula(
      `=IF(OR(A${row}="",J${row}=""),"",IF(G${row}="B",IF(J${row}="неуд","2F",IF(J${row}="уд","3E",IF(J${row}="хорошо","3D","3D/4C"))),IF(G${row}="A",IF(J${row}="неуд","2F",IF(J${row}="уд","3D",IF(J${row}="хорошо","4C","4B/5A"))),IF(G${row}="S",IF(J${row}="неуд","2F",IF(J${row}="уд","3D",IF(J${row}="хорошо","4B/5A","5A"))),""))))`
    );
  }
  SpreadsheetApp.flush();
  refreshRanking_(central);
  refreshPlusTotals_(central);
}

function krStatusFormula_(row, firstRow, lastRow) {
  const parts = ['E', 'F', 'G', 'H', 'I'].map(col =>
    `COUNTIFS('${DM.cw}'!$A$${firstRow}:$A$${lastRow},A${row},'${DM.cw}'!$${col}$${firstRow}:$${col}$${lastRow},">=3")`
  );
  return `=IF(A${row}="","",IF(${parts.join('+')}>=4,"зачет","незачет"))`;
}

function finalGradeFormula_(row) {
  return `=IF(A${row}="","",IF(AND(G${row}="зачет",H${row}="зачет"),IF(I${row}="F","2F",IF(L${row}="","",SWITCH(I${row}&"|"&L${row},"B|неуд","2F","B|уд","3E","B|хорошо","3D","B|очень хорошо",IF(M${row}="решена","4C","3D"),"A|неуд","2F","A|уд","3D","A|хорошо","4C","A|очень хорошо",IF(M${row}="решена","5A","4B"),"S|неуд","2F","S|уд","3D","S|хорошо","4B","S|очень хорошо","5A",""))),"2F"))`;
}

function refreshCentral_(central, plusFiles, sheetStates) {
  central = central || SpreadsheetApp.openById(getCentralId_());
  plusFiles = openPlusFiles_(plusFiles);
  SpreadsheetApp.flush();
  refreshLogCoefficients_(central, plusFiles, sheetStates);
  const common = central.getSheetByName(DM.common);
  const count = DM.commonLastRow - DM.commonFirstRow + 1;
  const mainFormulas = [];
  const gradeFormulas = [];
  for (let row = DM.commonFirstRow; row <= DM.commonLastRow; row++) {
    mainFormulas.push([
      `=IF(A${row}="","",SUMIF('${DM.logs}'!$A$${DM.logFirstRow}:$A$${DM.logMaxRow},A${row},'${DM.logs}'!$H$${DM.logFirstRow}:$H$${DM.logMaxRow}))`,
      `=IF(A${row}="","",SUMIF('${DM.cw}'!$A:$A,A${row},'${DM.cw}'!$D:$D))`,
      krStatusFormula_(row, 3, 92),
      krStatusFormula_(row, 96, 185),
      `=IF(A${row}="","",IF(AND(E${row}+IF(K${row}="",0,K${row})>=55,F${row}+IF(J${row}="",0,J${row})>=24),"S",IF(AND(E${row}+IF(K${row}="",0,K${row})>=35,F${row}+IF(J${row}="",0,J${row})>=14),"A",IF(E${row}+IF(K${row}="",0,K${row})>=17,"B","F"))))`
    ]);
    gradeFormulas.push([finalGradeFormula_(row)]);
  }
  common.getRange(DM.commonFirstRow, 5, count, 5).setFormulas(mainFormulas);
  common.getRange(DM.commonFirstRow, 14, count, 1).setFormulas(gradeFormulas);
  SpreadsheetApp.flush();
  refreshRanking_(central);
  refreshPlusTotals_(central, plusFiles, sheetStates);
}

function refreshRanking_(central) {
  const common = central.getSheetByName(DM.common);
  const ranking = central.getSheetByName(DM.ranking);
  const rows = common.getRange(DM.commonFirstRow, 1, 90, 9).getDisplayValues()
    .filter(row => String(row[0]).trim() && isActiveValue_(row[3]))
    .map(row => [row[0], row[1], row[4], row[5], row[8]]);
  const tierOrder = { 'S': 0, 'A': 1, 'B': 2, 'F': 3 };
  rows.sort((a, b) => {
    const tierA = tierOrder[a[4]] === undefined ? 99 : tierOrder[a[4]];
    const tierB = tierOrder[b[4]] === undefined ? 99 : tierOrder[b[4]];
    const tier = tierA - tierB;
    if (tier) return tier;
    const total = (toNumber_(b[2]) + toNumber_(b[3])) - (toNumber_(a[2]) + toNumber_(a[3]));
    if (total) return total;
    return String(a[0]).localeCompare(String(b[0]));
  });
  ranking.getRange(4, 1, 90, 8).clearContent();
  if (rows.length) ranking.getRange(4, 1, rows.length, 5).setValues(rows);
}

function syncGroups_(central, plusFiles) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    central = central || SpreadsheetApp.openById(getCentralId_());
    plusFiles = openPlusFiles_(plusFiles);
    const commonRows = central.getSheetByName(DM.common).getRange(DM.commonFirstRow, 1, 90, 4).getDisplayValues();
    const groups = [[], [], []];
    commonRows.forEach(row => {
      const practitionerNo = DM.practitioners.indexOf(String(row[1]).trim());
      const name = String(row[0]).trim();
      if (practitionerNo >= 0 && name) groups[practitionerNo].push({ name, tg: row[2] });
    });
    syncCwRoster_(central);
    const snapshots = collectSnapshots_(plusFiles);

    DM.practitioners.forEach((practitioner, index) => {
      const file = plusFiles[practitioner];
      if (!file) return;
      const layout = plusLayout_(practitioner);
      const roster = groups[index].slice(0, DM.studentsPerGroup);
      const taskCount = layout.taskLastCol - layout.taskFirstCol + 1;
      for (let practice = 1; practice <= DM.practiceCount; practice++) {
        const sheet = getPlusSheet_(file, practice);
        if (!sheet) continue;
        const first = DM.plusFirstRow;
        const names = Array.from({ length: DM.studentsPerGroup }, () => ['']);
        const presence = Array.from({ length: DM.studentsPerGroup }, () => ['']);
        const tasks = Array.from(
          { length: DM.studentsPerGroup },
          () => Array(taskCount).fill('')
        );
        roster.forEach((student, offset) => {
          names[offset][0] = student.name;
          const saved = snapshots.byName[student.name] && snapshots.byName[student.name][practice]
            ? snapshots.byName[student.name][practice]
            : snapshots.bySlot[practitioner] && snapshots.bySlot[practitioner][practice]
              ? snapshots.bySlot[practitioner][practice][offset]
              : null;
          if (saved) {
            if (layout.presenceCol) presence[offset][0] = saved.presence;
            const savedTasks = saved.tasks.slice(0, taskCount);
            while (savedTasks.length < taskCount) savedTasks.push('');
            tasks[offset] = savedTasks;
          }
        });
        sheet.getRange(first, 1, DM.studentsPerGroup, 1).setValues(names);
        if (layout.presenceCol) {
          sheet.getRange(first, layout.presenceCol, DM.studentsPerGroup, 1).setValues(presence);
        }
        sheet.getRange(first, layout.taskFirstCol, DM.studentsPerGroup, taskCount).setValues(tasks);
        const state = snapshots.sheets[practitioner][practice];
        state.names = names;
        state.presence = presence;
        state.tasks = tasks;
      }
    });
    clearPlusNotes_(plusFiles);
    reconcilePlusResults_(central, plusFiles, snapshots.sheets);
    syncVisibility_(central, plusFiles, snapshots.sheets);
    refreshCentral_(central, plusFiles, snapshots.sheets);
  } finally {
    lock.releaseLock();
  }
}

function syncCwRoster_(central) {
  const commonRows = central.getSheetByName(DM.common).getRange(DM.commonFirstRow, 1, 90, 2).getDisplayValues();
  const cw = central.getSheetByName(DM.cw);
  const saved = {};
  [3, 96].forEach(firstRow => {
    cw.getRange(firstRow, 1, 90, 9).getValues().forEach(row => {
      const name = String(row[0] == null ? '' : row[0]).trim();
      if (!name) return;
      saved[name] = { attempt: row[2], scores: row.slice(4, 9) };
    });
  });
  // Два блока по 90 студентов: шапки находятся в строках 1 и 95.
  [3, 96].forEach(firstRow => {
    const roster = [];
    const scores = [];
    commonRows.forEach(row => {
      const previous = saved[String(row[0]).trim()];
      roster.push([row[0], row[1], previous ? previous.attempt : '']);
      scores.push(previous ? previous.scores : ['', '', '', '', '']);
    });
    cw.getRange(firstRow, 1, commonRows.length, 3).setValues(roster);
    cw.getRange(firstRow, 5, commonRows.length, 5).setValues(scores);
  });
}

function collectSnapshots_(plusFiles) {
  const snapshots = { byName: {}, bySlot: {}, sheets: {} };
  plusFiles = openPlusFiles_(plusFiles);
  DM.practitioners.forEach(practitioner => {
    const file = plusFiles[practitioner];
    if (!file) return;
    const layout = plusLayout_(practitioner);
    snapshots.sheets[practitioner] = {};
    for (let practice = 1; practice <= DM.practiceCount; practice++) {
      const sheet = getPlusSheet_(file, practice);
      if (!sheet) continue;
      const first = DM.plusFirstRow;
      const block = sheet.getRange(
        DM.plusHeaderRow, 1, DM.studentsPerGroup + 1, layout.taskLastCol
      ).getValues();
      const rows = block.slice(first - DM.plusHeaderRow);
      snapshots.bySlot[practitioner] = snapshots.bySlot[practitioner] || {};
      snapshots.bySlot[practitioner][practice] = [];
      snapshots.sheets[practitioner][practice] = {
        sheet,
        headers: block[0].slice(layout.taskFirstCol - 1, layout.taskLastCol),
        names: rows.map(row => [row[0]]),
        presence: rows.map(row => [layout.presenceCol ? row[layout.presenceCol - 1] : '']),
        tasks: rows.map(row => row.slice(layout.taskFirstCol - 1, layout.taskLastCol))
      };
      rows.forEach(row => {
        const name = String(row[0] == null ? '' : row[0]).trim();
        snapshots.bySlot[practitioner][practice].push({
          presence: layout.presenceCol ? row[layout.presenceCol - 1] : '',
          tasks: row.slice(layout.taskFirstCol - 1, layout.taskLastCol)
        });
        if (!name) return;
        snapshots.byName[name] = snapshots.byName[name] || {};
        snapshots.byName[name][practice] = {
          presence: layout.presenceCol ? row[layout.presenceCol - 1] : '',
          tasks: row.slice(layout.taskFirstCol - 1, layout.taskLastCol)
        };
      });
    }
  });
  return snapshots;
}

function syncVisibility_(central, plusFiles, sheetStates) {
  central = central || SpreadsheetApp.openById(getCentralId_());
  plusFiles = openPlusFiles_(plusFiles);
  const rows = central.getSheetByName(DM.common).getRange(DM.commonFirstRow, 1, 90, 4).getDisplayValues();
  const active = {};
  rows.forEach(row => { if (row[0]) active[String(row[0]).trim()] = isActiveValue_(row[3]); });
  DM.practitioners.forEach(practitioner => {
    const file = plusFiles[practitioner];
    if (!file) return;
    for (let practice = 1; practice <= DM.practiceCount; practice++) {
      const sheet = getPlusSheet_(file, practice);
      if (!sheet) continue;
      const first = DM.plusFirstRow;
      const state = sheetStates && sheetStates[practitioner]
        ? sheetStates[practitioner][practice] : null;
      const names = state
        ? state.names
        : sheet.getRange(first, 1, DM.studentsPerGroup, 1).getDisplayValues();
      sheet.showRows(first, DM.studentsPerGroup);
      let hiddenStart = -1;
      for (let offset = 0; offset <= names.length; offset++) {
        const name = offset < names.length ? String(names[offset][0]).trim() : '';
        const hide = offset < names.length && name !== '' && active[name] === false;
        if (hide && hiddenStart < 0) hiddenStart = offset;
        if (!hide && hiddenStart >= 0) {
          sheet.hideRows(first + hiddenStart, offset - hiddenStart);
          hiddenStart = -1;
        }
      }
    }
  });
}

function refreshPlusTotals_(central, plusFiles, sheetStates) {
  plusFiles = openPlusFiles_(plusFiles);
  const logs = central.getSheetByName(DM.logs);
  const lastRow = findLastLogRow_(logs);
  const totals = {};
  if (lastRow >= DM.logFirstRow) {
    logs.getRange(DM.logFirstRow, 1, lastRow - DM.logFirstRow + 1, 9).getDisplayValues().forEach(row => {
      const student = String(row[0]).trim();
      const practice = Number(row[2]);
      const points = toNumber_(row[7]);
      if (student && practice && points) totals[student + '|' + practice] = (totals[student + '|' + practice] || 0) + points;
    });
  }
  DM.practitioners.forEach(practitioner => {
    const file = plusFiles[practitioner];
    if (!file) return;
    const layout = plusLayout_(practitioner);
    for (let practice = 1; practice <= DM.practiceCount; practice++) {
      const sheet = getPlusSheet_(file, practice);
      if (!sheet) continue;
      const first = DM.plusFirstRow;
      const state = sheetStates && sheetStates[practitioner]
        ? sheetStates[practitioner][practice] : null;
      const names = state
        ? state.names
        : sheet.getRange(first, 1, DM.studentsPerGroup, 1).getDisplayValues();
      const values = names.map(row => [row[0] ? (totals[row[0].trim() + '|' + practice] || 0) : '']);
      sheet.getRange(first, layout.totalCol, DM.studentsPerGroup, 1).setValues(values);
    }
  });
}

function clearPlusNotes_(plusFiles) {
  plusFiles = openPlusFiles_(plusFiles);
  DM.practitioners.forEach(practitioner => {
    const file = plusFiles[practitioner];
    if (!file) return;
    const layout = plusLayout_(practitioner);
    for (let practice = 1; practice <= DM.practiceCount; practice++) {
      const sheet = getPlusSheet_(file, practice);
      if (!sheet) continue;
      sheet.getRange(DM.plusFirstRow, layout.taskFirstCol,
        DM.studentsPerGroup, layout.taskLastCol - layout.taskFirstCol + 1).clearNote();
    }
  });
}

function refreshLogCoefficients_(central, plusFiles, sheetStates) {
  const coefficients = {};
  plusFiles = openPlusFiles_(plusFiles);
  DM.practitioners.forEach(practitioner => {
    const file = plusFiles[practitioner];
    if (!file) return;
    const layout = plusLayout_(practitioner);
    for (let practice = 1; practice <= DM.practiceCount; practice++) {
      const sheet = getPlusSheet_(file, practice);
      if (!sheet) continue;
      const state = sheetStates && sheetStates[practitioner]
        ? sheetStates[practitioner][practice] : null;
      if (state) {
        const values = sheet.getRange(
          DM.plusFirstRow, layout.coefficientCol, DM.studentsPerGroup, 1
        ).getValues();
        state.names.forEach((row, index) => {
          const name = String(row[0] == null ? '' : row[0]).trim();
          if (name) coefficients[name + '|' + practice] = Number(values[index][0]) || 0.5;
        });
      } else {
        sheet.getRange(DM.plusFirstRow, 1, DM.studentsPerGroup, layout.coefficientCol)
          .getValues().forEach(row => {
            const name = String(row[0] == null ? '' : row[0]).trim();
            if (name) coefficients[name + '|' + practice] = Number(row[layout.coefficientCol - 1]) || 0.5;
          });
      }
    }
  });

  const logs = central.getSheetByName(DM.logs);
  const lastRow = findLastLogRow_(logs);
  if (lastRow < DM.logFirstRow) return;
  const count = lastRow - DM.logFirstRow + 1;
  const rows = logs.getRange(DM.logFirstRow, 1, count, 9).getValues();
  const resultFormulas = logs.getRange(DM.logFirstRow, 8, count, 1).getFormulas();
  const coefficientAndResult = rows.map((row, index) => [
    row[6],
    resultFormulas[index][0] || row[7]
  ]);
  rows.forEach((row, index) => {
    const base = row[5];
    const student = String(row[0] == null ? '' : row[0]).trim();
    const practice = Number(row[2]);
    const key = student + '|' + practice;
    const isNormal = student && practice && base !== '' && base != null && coefficients[key] != null;
    if (!isNormal) return;
    coefficientAndResult[index][0] = coefficients[key];
    const targetRow = DM.logFirstRow + index;
    coefficientAndResult[index][1] =
      `=IF(OR(F${targetRow}="",G${targetRow}=""),"",F${targetRow}*G${targetRow}+IF(J${targetRow}="",0,J${targetRow}))`;
  });
  logs.getRange(DM.logFirstRow, 7, count, 2).setValues(coefficientAndResult);
}

function reconcilePlusResults_(central, plusFiles, sheetStates) {
  plusFiles = openPlusFiles_(plusFiles);
  const logs = central.getSheetByName(DM.logs);
  const markers = {};
  const markerOrder = [];

  DM.practitioners.forEach((practitioner, practitionerIndex) => {
    const file = plusFiles[practitioner];
    if (!file) return;
    const layout = plusLayout_(practitioner);
    for (let practice = 1; practice <= DM.practiceCount; practice++) {
      const sheet = getPlusSheet_(file, practice);
      if (!sheet) continue;
      const state = sheetStates && sheetStates[practitioner]
        ? sheetStates[practitioner][practice] : null;
      let headers;
      let names;
      let taskRows;
      if (state) {
        headers = state.headers;
        names = state.names;
        taskRows = state.tasks;
      } else {
        const block = sheet.getRange(
          DM.plusHeaderRow, 1, DM.studentsPerGroup + 1, layout.taskLastCol
        ).getDisplayValues();
        headers = block[0].slice(layout.taskFirstCol - 1, layout.taskLastCol);
        const rows = block.slice(1);
        names = rows.map(row => [row[0]]);
        taskRows = rows.map(row => row.slice(layout.taskFirstCol - 1, layout.taskLastCol));
      }
      names.forEach((row, studentIndex) => {
        const student = String(row[0]).trim();
        if (!student) return;
        headers.forEach((task, offset) => {
          const value = String(taskRows[studentIndex][offset]).trim();
          if (value !== '1' && value !== '-') return;
          const key = logIdentityKey_(student, practice, task);
          if (markers[key]) return;
          const marker = {
            key,
            kind: value === '1' ? 'normal' : 'penalty',
            file,
            sheet,
            practitionerNo: practitionerIndex + 1,
            practitioner,
            practice,
            student,
            task,
            hasLog: false
          };
          markers[key] = marker;
          markerOrder.push(marker);
        });
      });
    }
  });

  // Состояние 1/- в плюсовиках является источником истины. Оставляем ровно
  // одну соответствующую запись, а старые, противоположные и дубли удаляем.
  const lastRow = findLastLogRow_(logs);
  const rangesToClear = [];
  if (lastRow >= DM.logFirstRow) {
    logs.getRange(DM.logFirstRow, 1, lastRow - DM.logFirstRow + 1, 9)
      .getDisplayValues().forEach((row, index) => {
        const student = String(row[0]).trim();
        const practice = Number(row[2]);
        const task = String(row[3]).trim();
        const base = String(row[5]).trim();
        const coefficient = String(row[6]).trim();
        const result = toNumber_(row[7]);
        const isPenalty = base === '' && coefficient === '' && result === -2;
        const isNormal = student !== '' && practice > 0 && task !== '' && base !== '';
        if (!isPenalty && !isNormal) return;

        const marker = markers[logIdentityKey_(student, practice, task)];
        const kind = isPenalty ? 'penalty' : 'normal';
        if (marker && marker.kind === kind && !marker.hasLog) {
          marker.hasLog = true;
          return;
        }
        const sheetRow = DM.logFirstRow + index;
        rangesToClear.push(`A${sheetRow}:K${sheetRow}`);
      });
  }
  if (rangesToClear.length) {
    logs.getRangeList(rangesToClear).clearContent();
    SpreadsheetApp.flush();
  }

  markerOrder.forEach(marker => {
    if (marker.hasLog) return;
    if (marker.kind === 'normal') {
      const appendResult = appendNormalLog_(central, marker.file, marker.sheet,
        marker.practitionerNo, marker.practitioner, marker.practice,
        marker.student, marker.task);
      marker.hasLog = appendResult.ok;
      return;
    }
    appendPenaltyLog_(central, marker.practitioner, marker.practice,
      marker.student, marker.task);
    marker.hasLog = true;
  });
}

function practiceFromSheet_(sheet) {
  const match = String(sheet.getName()).match(/^Практика\s+(\d+)$/i);
  if (!match) return 0;
  const practice = Number(match[1]);
  return practice >= 1 && practice <= DM.practiceCount ? practice : 0;
}

function getPlusSheet_(file, practice) {
  return file.getSheetByName(`Практика ${practice}`);
}

function findStudentRow_(sheet, name, first, last) {
  const values = sheet.getRange(first, 1, last - first + 1, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) if (String(values[i][0]).trim() === name) return first + i;
  return 0;
}

function findLastLogRow_(sheet) {
  const values = sheet.getRange(DM.logFirstRow, 1, DM.logMaxRow - DM.logFirstRow + 1, 1).getDisplayValues();
  let last = DM.logFirstRow - 1;
  values.forEach((row, index) => { if (String(row[0]).trim()) last = DM.logFirstRow + index; });
  return last;
}

function isStudentActive_(central, practitioner, student) {
  const rows = central.getSheetByName(DM.common).getRange(DM.commonFirstRow, 1, 90, 4).getDisplayValues();
  return rows.some(row => String(row[0]).trim() === student && String(row[1]).trim() === practitioner && isActiveValue_(row[3]));
}

function isActiveValue_(value) {
  const text = String(value == null ? '' : value).trim().toLowerCase();
  return !text || (text !== 'нет' && text !== 'no' && text !== 'n');
}

function openPlusFiles_(knownFiles) {
  const files = knownFiles || {};
  DM.practitioners.forEach(practitioner => {
    const fileId = DM.plusFileIds[practitioner];
    if (!files[practitioner] && fileId) {
      files[practitioner] = SpreadsheetApp.openById(fileId);
    }
  });
  return files;
}

function setSystemTimeZones_(central, plusFiles) {
  plusFiles = openPlusFiles_(plusFiles);
  const files = [central].concat(
    DM.practitioners.map(practitioner => plusFiles[practitioner]).filter(Boolean)
  );
  files.forEach(file => {
    if (file.getSpreadsheetTimeZone() !== DM.timeZone) {
      file.setSpreadsheetTimeZone(DM.timeZone);
    }
  });

  const centralLogs = central.getSheetByName(DM.logs);
  if (centralLogs) {
    centralLogs.getRange(
      DM.logFirstRow, 11, DM.logMaxRow - DM.logFirstRow + 1, 1
    ).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
  DM.practitioners.forEach(practitioner => {
    const file = plusFiles[practitioner];
    const localLogs = file && file.getSheetByName(DM.plusLogs);
    if (!localLogs || localLogs.getMaxRows() < 2) return;
    localLogs.getRange(2, 5, localLogs.getMaxRows() - 1, 1)
      .setNumberFormat('yyyy-mm-dd hh:mm:ss');
  });
}

function getCentralId_() {
  return DM.centralFileId || SpreadsheetApp.getActive().getId();
}

function validateConfig_() {
  const missing = [];
  if (!DM.centralFileId) missing.push('centralFileId');
  DM.practitioners.forEach(name => {
    if (!DM.plusFileIds[name]) missing.push(`plusFileIds[${name}]`);
  });
  if (missing.length) {
    throw new Error(
      'Заполните ID четырёх сконвертированных Google-таблиц в начале Code.gs: ' +
      missing.join(', ')
    );
  }
}

function ladderPoints_(exitNo) {
  if (exitNo <= 3) return 6;
  if (exitNo <= 6) return 5;
  if (exitNo <= 9) return 4;
  if (exitNo <= 12) return 3;
  if (exitNo <= 15) return 2;
  return 0;
}

function tasksEqual_(left, right) {
  const a = String(left == null ? '' : left).trim();
  const b = String(right == null ? '' : right).trim();
  if (a === b) return true;
  const na = Number(a.replace(',', '.'));
  const nb = Number(b.replace(',', '.'));
  return a !== '' && b !== '' && !isNaN(na) && !isNaN(nb) && na === nb;
}

function logIdentityKey_(student, practice, task) {
  const taskText = String(task == null ? '' : task).trim();
  const taskNumber = Number(taskText.replace(',', '.'));
  const normalizedTask = taskText !== '' && !isNaN(taskNumber)
    ? String(taskNumber) : taskText;
  return [
    String(student == null ? '' : student).trim(),
    Number(practice) || String(practice == null ? '' : practice).trim(),
    normalizedTask
  ].join('|');
}

function toNumber_(value) {
  const n = Number(String(value == null ? '' : value).replace(',', '.'));
  return isNaN(n) ? 0 : n;
}
