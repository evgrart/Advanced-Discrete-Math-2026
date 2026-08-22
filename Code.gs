const DM = {
  // Вставьте ID четырёх Google-таблиц. ID находится между /d/ и /edit в URL.
  centralFileId: '1RIDAw8VWB76L6E2RDaiSiyj9UlLk0c0P-rfs3WivcFM',
  plusFileIds: {
    'Артём': '',
    'Рами': '1WG-5ZdLOjKTiy9hG-muEugFT16jjxcCQMi3_N46rA3c',
    'Немат': '1SnTIZRF7mJByn0WEZyC1zooNE3y1gG7r2TXWos1HW_o'
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
  logMaxRow: 3006
};

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
  syncGroups_();
  refreshCentral_();
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
  if (col < DM.taskFirstCol || col > DM.taskLastCol) return;
  cell.clearNote();

  const practiceNo = practiceFromSheet_(sheet);
  if (!practiceNo) return;
  const plusHeader = DM.plusHeaderRow;
  const plusFirst = DM.plusFirstRow;
  const plusLast = plusFirst + DM.studentsPerGroup - 1;
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
    refreshCentral_();
    return;
  }

  if (value === '-') {
    removeLogRecord_(central, practitioner, practiceNo, studentName, taskName, false);
    if (oldValue !== '-') appendPenaltyLog_(central, practitioner, practiceNo, studentName, taskName);
    refreshCentral_();
    return;
  }

  if (value === '' || value === '+') {
    if (oldValue === '1' || oldValue === '-') {
      removeLogRecord_(central, practitioner, practiceNo, studentName, taskName, false);
    }
    refreshCentral_();
  }
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
  const coefficient = Number(plusSheet.getRange(studentRow, DM.plusCoefficientCol).getValue()) || 0.5;
  const exitNo = studentExit + 1;
  const base = ladderPoints_(exitNo);
  const newRow = Math.max(findLastLogRow_(logs) + 1, DM.logFirstRow);
  logs.getRange(newRow, 1, 1, 11).setValues([[
    studentName, practitioner, practiceNo, taskName, exitNo,
    base, coefficient, '', exitNo > 15 ? 'Дополнительный выход' : 'Назначен', '', new Date()
  ]]);
  logs.getRange(newRow, 8).setFormula(
    `=IF(OR(F${newRow}="",G${newRow}=""),"",F${newRow}*G${newRow}+IF(J${newRow}="",0,J${newRow}))`
  );
  logs.getRange(newRow, 11).setNumberFormat('yyyy-mm-dd hh:mm');
  return { ok: true, message: `Начислено D: ${base * coefficient}` };
}

function appendPlusHistory_(plusFile, studentName, practiceNo, taskName, state) {
  const logs = plusFile.getSheetByName(DM.plusLogs);
  if (!logs) return;
  const lastRow = Math.max(logs.getLastRow() + 1, 2);
  logs.getRange(lastRow, 1, 1, 5).setValues([[
    studentName, practiceNo, taskName, state, new Date()
  ]]);
  logs.getRange(lastRow, 5).setNumberFormat('yyyy-mm-dd hh:mm');
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
  logs.getRange(newRow, 11).setNumberFormat('yyyy-mm-dd hh:mm');
}

function removeLogRecord_(central, practitioner, practiceNo, studentName, taskName, onlyPenalty) {
  const logs = central.getSheetByName(DM.logs);
  const lastRow = findLastLogRow_(logs);
  if (lastRow < DM.logFirstRow) return;
  const rows = logs.getRange(DM.logFirstRow, 1, lastRow - DM.logFirstRow + 1, 9).getDisplayValues();
  rows.forEach((row, index) => {
    const penalty = String(row[5]).trim() === '' && String(row[6]).trim() === '' && Number(row[7]) === -2;
    const same = String(row[0]).trim() === studentName &&
      String(row[1]).trim() === practitioner && Number(row[2]) === practiceNo &&
      tasksEqual_(row[3], taskName) && (!onlyPenalty || penalty);
    if (same) logs.getRange(DM.logFirstRow + index, 1, 1, 11).clearContent();
  });
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
  logs.getRange(newRow, 11).setNumberFormat('yyyy-mm-dd hh:mm');
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
  clearPlusNotes_();
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

function refreshCentral_() {
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
    common.getRange(row, 7).setFormula(krStatusFormula_(row, 3, 92));
    common.getRange(row, 8).setFormula(krStatusFormula_(row, 96, 185));
    common.getRange(row, 9).setFormula(
      `=IF(A${row}="","",IF(AND(E${row}+IF(K${row}="",0,K${row})>=55,F${row}+IF(J${row}="",0,J${row})>=24),"S",IF(AND(E${row}+IF(K${row}="",0,K${row})>=35,F${row}+IF(J${row}="",0,J${row})>=14),"A",IF(E${row}+IF(K${row}="",0,K${row})>=17,"B","F"))))`
    );
    common.getRange(row, 14).setFormula(finalGradeFormula_(row));
  }
  SpreadsheetApp.flush();
  refreshRanking_(central);
  refreshPlusTotals_(central);
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

function syncGroups_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const central = SpreadsheetApp.openById(getCentralId_());
    const commonRows = central.getSheetByName(DM.common).getRange(DM.commonFirstRow, 1, 90, 4).getDisplayValues();
    const groups = [[], [], []];
    commonRows.forEach(row => {
      const practitionerNo = DM.practitioners.indexOf(String(row[1]).trim());
      const name = String(row[0]).trim();
      if (practitionerNo >= 0 && name) groups[practitionerNo].push({ name, tg: row[2] });
    });
    syncCwRoster_(central);
    const snapshots = collectSnapshots_();

    DM.practitioners.forEach((practitioner, index) => {
      const fileId = DM.plusFileIds[practitioner];
      if (!fileId) return;
      const file = SpreadsheetApp.openById(fileId);
      for (let practice = 1; practice <= DM.practiceCount; practice++) {
        const sheet = getPlusSheet_(file, practice);
        if (!sheet) continue;
        const first = DM.plusFirstRow;
        sheet.getRange(first, 1, DM.studentsPerGroup, 1).clearContent();
        sheet.getRange(first, 2, DM.studentsPerGroup, 1).clearContent();
        sheet.getRange(first, DM.taskFirstCol, DM.studentsPerGroup, 30).clearContent();
        groups[index].slice(0, DM.studentsPerGroup).forEach((student, offset) => {
          const row = first + offset;
          sheet.getRange(row, 1).setValue(student.name);
          const saved = snapshots.byName[student.name] && snapshots.byName[student.name][practice]
            ? snapshots.byName[student.name][practice]
            : snapshots.bySlot[practitioner] && snapshots.bySlot[practitioner][practice]
              ? snapshots.bySlot[practitioner][practice][offset]
              : null;
          if (saved) {
            sheet.getRange(row, 2).setValue(saved.presence);
            sheet.getRange(row, DM.taskFirstCol, 1, 30).setValues([saved.tasks]);
          }
        });
      }
    });
    clearPlusNotes_();
    reconcilePlusResults_(central);
    syncVisibility_();
    refreshCentral_();
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
    commonRows.forEach((row, offset) => {
      const target = firstRow + offset;
      const previous = saved[String(row[0]).trim()];
      cw.getRange(target, 1, 1, 2).setValues([[row[0], row[1]]]);
      cw.getRange(target, 3).setValue(previous ? previous.attempt : '');
      cw.getRange(target, 5, 1, 5).setValues([previous ? previous.scores : ['', '', '', '', '']]);
    });
  });
}

function collectSnapshots_() {
  const snapshots = { byName: {}, bySlot: {} };
  DM.practitioners.forEach(practitioner => {
    const fileId = DM.plusFileIds[practitioner];
    if (!fileId) return;
    const file = SpreadsheetApp.openById(fileId);
    for (let practice = 1; practice <= DM.practiceCount; practice++) {
      const sheet = getPlusSheet_(file, practice);
      if (!sheet) continue;
      const first = DM.plusFirstRow;
      const rows = sheet.getRange(first, 1, DM.studentsPerGroup, 35).getValues();
      snapshots.bySlot[practitioner] = snapshots.bySlot[practitioner] || {};
      snapshots.bySlot[practitioner][practice] = [];
      rows.forEach(row => {
        const name = String(row[0] == null ? '' : row[0]).trim();
        snapshots.bySlot[practitioner][practice].push({
          presence: row[1],
          tasks: row.slice(5, 35)
        });
        if (!name) return;
        snapshots.byName[name] = snapshots.byName[name] || {};
        snapshots.byName[name][practice] = {
          presence: row[1],
          tasks: row.slice(5, 35)
        };
      });
    }
  });
  return snapshots;
}

function syncVisibility_() {
  const central = SpreadsheetApp.openById(getCentralId_());
  const rows = central.getSheetByName(DM.common).getRange(DM.commonFirstRow, 1, 90, 4).getDisplayValues();
  const active = {};
  rows.forEach(row => { if (row[0]) active[String(row[0]).trim()] = isActiveValue_(row[3]); });
  DM.practitioners.forEach(practitioner => {
    const fileId = DM.plusFileIds[practitioner];
    if (!fileId) return;
    const file = SpreadsheetApp.openById(fileId);
    for (let practice = 1; practice <= DM.practiceCount; practice++) {
      const sheet = getPlusSheet_(file, practice);
      if (!sheet) continue;
      const first = DM.plusFirstRow;
      const names = sheet.getRange(first, 1, DM.studentsPerGroup, 1).getDisplayValues();
      names.forEach((entry, offset) => {
        const show = entry[0] === '' || active[String(entry[0]).trim()] !== false;
        if (show) sheet.showRows(first + offset);
        else sheet.hideRows(first + offset);
      });
    }
  });
}

function refreshPlusTotals_(central) {
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
    const fileId = DM.plusFileIds[practitioner];
    if (!fileId) return;
    const file = SpreadsheetApp.openById(fileId);
    for (let practice = 1; practice <= DM.practiceCount; practice++) {
      const sheet = getPlusSheet_(file, practice);
      if (!sheet) continue;
      const first = DM.plusFirstRow;
      const names = sheet.getRange(first, 1, DM.studentsPerGroup, 1).getDisplayValues();
      const values = names.map(row => [row[0] ? (totals[row[0].trim() + '|' + practice] || 0) : '']);
      sheet.getRange(first, DM.plusTotalCol, DM.studentsPerGroup, 1).setValues(values);
    }
  });
}

function clearPlusNotes_() {
  DM.practitioners.forEach(practitioner => {
    const fileId = DM.plusFileIds[practitioner];
    if (!fileId) return;
    const file = SpreadsheetApp.openById(fileId);
    for (let practice = 1; practice <= DM.practiceCount; practice++) {
      const sheet = getPlusSheet_(file, practice);
      if (!sheet) continue;
      sheet.getRange(DM.plusFirstRow, DM.taskFirstCol,
        DM.studentsPerGroup, DM.taskLastCol - DM.taskFirstCol + 1).clearNote();
    }
  });
}

function refreshLogCoefficients_(central) {
  const coefficients = {};
  DM.practitioners.forEach(practitioner => {
    const fileId = DM.plusFileIds[practitioner];
    if (!fileId) return;
    const file = SpreadsheetApp.openById(fileId);
    for (let practice = 1; practice <= DM.practiceCount; practice++) {
      const sheet = getPlusSheet_(file, practice);
      if (!sheet) continue;
      sheet.getRange(DM.plusFirstRow, 1, DM.studentsPerGroup, DM.plusCoefficientCol)
        .getValues().forEach(row => {
          const name = String(row[0] == null ? '' : row[0]).trim();
          if (name) coefficients[name + '|' + practice] = Number(row[DM.plusCoefficientCol - 1]) || 0.5;
        });
    }
  });

  const logs = central.getSheetByName(DM.logs);
  const lastRow = findLastLogRow_(logs);
  if (lastRow < DM.logFirstRow) return;
  const count = lastRow - DM.logFirstRow + 1;
  const rows = logs.getRange(DM.logFirstRow, 1, count, 9).getValues();
  const coefficientValues = rows.map(row => [row[6]]);
  rows.forEach((row, index) => {
    const base = row[5];
    const student = String(row[0] == null ? '' : row[0]).trim();
    const practice = Number(row[2]);
    const key = student + '|' + practice;
    const isNormal = student && practice && base !== '' && base != null && coefficients[key] != null;
    if (!isNormal) return;
    coefficientValues[index][0] = coefficients[key];
    const targetRow = DM.logFirstRow + index;
    logs.getRange(targetRow, 8).setFormula(
      `=IF(OR(F${targetRow}="",G${targetRow}=""),"",F${targetRow}*G${targetRow}+IF(J${targetRow}="",0,J${targetRow}))`
    );
  });
  logs.getRange(DM.logFirstRow, 7, count, 1).setValues(coefficientValues);
}

function reconcilePlusResults_(central) {
  const logs = central.getSheetByName(DM.logs);
  const lastRow = findLastLogRow_(logs);
  const normal = {};
  const penalties = {};
  if (lastRow >= DM.logFirstRow) {
    logs.getRange(DM.logFirstRow, 1, lastRow - DM.logFirstRow + 1, 9)
      .getDisplayValues().forEach(row => {
        const key = [row[0], row[1], row[2], row[3]].map(value => String(value).trim()).join('|');
        const isPenalty = String(row[5]).trim() === '' && String(row[6]).trim() === '' && Number(row[7]) === -2;
        if (isPenalty) penalties[key] = true;
        else if (row[0] && row[1] && row[2] && row[3]) normal[key] = true;
      });
  }

  DM.practitioners.forEach((practitioner, practitionerIndex) => {
    const fileId = DM.plusFileIds[practitioner];
    if (!fileId) return;
    const file = SpreadsheetApp.openById(fileId);
    for (let practice = 1; practice <= DM.practiceCount; practice++) {
      const sheet = getPlusSheet_(file, practice);
      if (!sheet) continue;
      const headers = sheet.getRange(DM.plusHeaderRow, DM.taskFirstCol, 1,
        DM.taskLastCol - DM.taskFirstCol + 1).getDisplayValues()[0];
      const rows = sheet.getRange(DM.plusFirstRow, 1, DM.studentsPerGroup,
        DM.taskLastCol).getDisplayValues();
      rows.forEach(row => {
        const student = String(row[0]).trim();
        if (!student) return;
        headers.forEach((task, offset) => {
          const value = String(row[DM.taskFirstCol - 1 + offset]).trim();
          if (value !== '1' && value !== '-') return;
          const key = [student, practitioner, practice, task].join('|');
          if (value === '1') {
            if (penalties[key]) {
              removeLogRecord_(central, practitioner, practice, student, task, false);
              delete penalties[key];
            }
            if (!normal[key]) {
              const result = appendNormalLog_(central, file, sheet, practitionerIndex + 1,
                practitioner, practice, student, task);
              if (result.ok) normal[key] = true;
            }
          } else {
            if (normal[key]) {
              removeLogRecord_(central, practitioner, practice, student, task, false);
              delete normal[key];
            }
            if (!penalties[key]) {
              appendPenaltyLog_(central, practitioner, practice, student, task);
              penalties[key] = true;
            }
          }
        });
      });
    }
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

function toNumber_(value) {
  const n = Number(String(value == null ? '' : value).replace(',', '.'));
  return isNaN(n) ? 0 : n;
}
