const DM = {
  common: 'Общий',
  logs: 'Логи',
  ranking: 'Рейтинг',
  plusSheets: ['Плюсы_Артём', 'Плюсы_Рами', 'Плюсы_Немат'],
  practitioners: ['Артём', 'Рами', 'Немат'],
  practiceCount: 15,
  blockSize: 33,
  studentsPerGroup: 30,
  taskFirstCol: 6,
  taskLastCol: 35,
  commonFirstRow: 4,
  commonLastRow: 93,
  logFirstRow: 4,
  logMaxRow: 3006,
  // Заполните адресами практиков, если нужно технически запретить студентам 1 и -.
  // Пустой список оставлен для совместимости с тестовыми аккаунтами.
  practitionerEmails: []
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('DM система')
    .addItem('Установить проверки и цвета', 'installDmSystem')
    .addItem('Синхронизировать группы', 'syncGroupAssignments')
    .addItem('Синхронизировать активность', 'syncActiveStudents')
    .addItem('Пересчитать D, CW и тиры', 'refreshTotals')
    .addToUi();
}

function installDmSystem() {
  setupValidations_();
  applyConditionalFormatting_();
  syncGroupAssignments();
  refreshTotals();
  SpreadsheetApp.getActive().toast('Проверки, цвета, группы и расчёты обновлены.');
}

function onEdit(e) {
  if (!e || !e.range) return;
  const range = e.range;
  const sheet = range.getSheet();
  const values = range.getValues();
  const row0 = range.getRow();
  const col0 = range.getColumn();

  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const cell = sheet.getRange(row0 + r, col0 + c);
      const row = cell.getRow();
      const col = cell.getColumn();

      if (sheet.getName() === DM.common && row >= DM.commonFirstRow && row <= DM.commonLastRow) {
        if (col >= 1 && col <= 3) {
          if (col === 2) handlePractitionerEdit_(row, e);
          syncGroupAssignments();
        } else if (col === 4) {
          handleПродEdit_(row);
          syncVisibility_();
          refreshTotals();
        }
        continue;
      }

      if (sheet.getName() === 'CW' ||
          (sheet.getName() === DM.logs && row >= DM.logFirstRow && (col === 8 || col === 10))) {
        refreshTotals();
        continue;
      }

      if (DM.plusSheets.indexOf(sheet.getName()) !== -1) {
        processPlusEdit_(sheet, cell, e);
      }
    }
  }
}

function processPlusEdit_(sheet, cell, event) {
  const col = cell.getColumn();
  const row = cell.getRow();
  if (col < DM.taskFirstCol || col > DM.taskLastCol) return;

  const practitionerNo = DM.plusSheets.indexOf(sheet.getName()) + 1;
  for (let practiceNo = 1; practiceNo <= DM.practiceCount; practiceNo++) {
    const blockStart = 4 + (practiceNo - 1) * DM.blockSize;
    const plusHeader = blockStart + 1;
    const plusFirst = blockStart + 2;
    const plusLast = plusFirst + DM.studentsPerGroup - 1;
    if (row < plusFirst || row > plusLast) continue;

    const value = String(cell.getDisplayValue()).trim();
    const oldValue = event && event.range.getNumRows() === 1 && event.range.getNumColumns() === 1
      ? String(event.oldValue || '').trim()
      : '';
    const studentName = String(sheet.getRange(row, 1).getDisplayValue()).trim();
    const taskName = String(sheet.getRange(plusHeader, col).getDisplayValue()).trim();
    if (!studentName || !taskName) return;

    if ((value === '1' || value === '-') && !isPractitionerEditor_()) {
      cell.setValue(oldValue === '+' ? '+' : '');
      SpreadsheetApp.getActive().toast('1 и - доступны только практикам.');
      return;
    }

    if (value === '1') {
      removeLogRecord_(practitionerNo, practiceNo, studentName, taskName, 'Штраф -2');
      if (oldValue !== '1') {
        const result = registerUpperOne_(practitionerNo, practiceNo, studentName, taskName);
        cell.setNote(result.message || '');
        if (!result.ok) {
          cell.setValue(oldValue === '+' ? '+' : '');
          return;
        }
      }
      refreshTotals();
    } else if (value === '-') {
      removeLogRecord_(practitionerNo, practiceNo, studentName, taskName);
      if (oldValue !== '-') registerPenalty_(practitionerNo, practiceNo, studentName, taskName);
      cell.setNote('Штраф -2 балла без коэффициента.');
      refreshTotals();
    } else if (value === '' || value === '+') {
      if (oldValue === '1' || oldValue === '-') {
        removeLogRecord_(practitionerNo, practiceNo, studentName, taskName);
        refreshTotals();
      }
    }
    return;
  }
}

function registerUpperOne_(practitionerNo, practiceNo, studentName, taskName) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    if (!isStudentActive_(practitionerNo, studentName)) {
      return { ok: false, message: 'Студент неактивен на листе «Общий».' };
    }
    const ss = SpreadsheetApp.getActive();
    const plusSheet = ss.getSheetByName(DM.plusSheets[practitionerNo - 1]);
    return appendLogRow_(ss.getSheetByName(DM.logs), plusSheet,
      practitionerNo, practiceNo, studentName, taskName);
  } finally {
    lock.releaseLock();
  }
}

function appendLogRow_(logs, plusSheet, practitionerNo, practiceNo, studentName, taskName) {
  const lastRow = findLastLogRow_(logs);
  const data = lastRow >= DM.logFirstRow
    ? logs.getRange(DM.logFirstRow, 1, lastRow - DM.logFirstRow + 1, 9).getDisplayValues()
    : [];
  const practitionerName = DM.practitioners[practitionerNo - 1];
  let studentExit = 0;
  let pairCount = 0;
  let duplicateCount = 0;

  data.forEach(row => {
    const rowStudent = String(row[0]).trim();
    const rowStatus = String(row[8]).trim();
    const isPenalty = String(row[5]).trim() === '' && String(row[6]).trim() === '' && Number(row[7]) === -2;
    if (rowStudent !== studentName || isPenalty || rowStatus === 'Ушёл на базу' || rowStatus === 'ушел на базу') return;
    studentExit++;
    if (Number(row[2]) === practiceNo) pairCount++;
    if (Number(row[2]) === practiceNo && tasksEqual_(row[3], taskName)) duplicateCount++;
  });

  if (pairCount >= 2) return { ok: false, message: 'У студента уже два выхода за эту практику.' };
  if (duplicateCount > 0) return { ok: false, message: 'Эта задача уже есть в «Логах».' };

  const blockStart = 4 + (practiceNo - 1) * DM.blockSize;
  const studentRow = findStudentRow_(plusSheet, studentName, blockStart + 2, blockStart + 31);
  if (!studentRow) return { ok: false, message: 'Студент не найден на плюсовике.' };
  const coefficient = Number(plusSheet.getRange(studentRow, 4).getValue()) || 0.5;
  const exitNo = studentExit + 1;
  const base = ladderPoints_(exitNo);
  const newRow = Math.max(findLastLogRow_(logs) + 1, DM.logFirstRow);

  logs.getRange(newRow, 1, 1, 11).setValues([[
    studentName, practitionerName, practiceNo, taskName, exitNo,
    base, coefficient, '', exitNo > 15 ? 'Дополнительный выход' : 'Назначен', '', new Date()
  ]]);
  logs.getRange(newRow, 8).setFormula(`=IF(OR(F${newRow}="",G${newRow}=""),"",F${newRow}*G${newRow}+IF(J${newRow}="",0,J${newRow}))`);
  logs.getRange(newRow, 11).setNumberFormat('yyyy-mm-dd hh:mm');
  return { ok: true, message: `Начислено D: ${base * coefficient}` };
}

function registerPenalty_(practitionerNo, practiceNo, studentName, taskName) {
  const logs = SpreadsheetApp.getActive().getSheetByName(DM.logs);
  const newRow = Math.max(findLastLogRow_(logs) + 1, DM.logFirstRow);
  logs.getRange(newRow, 1, 1, 11).setValues([[
    studentName, DM.practitioners[practitionerNo - 1], practiceNo, taskName,
    '', '', '', -2, '', '', new Date()
  ]]);
  logs.getRange(newRow, 11).setNumberFormat('yyyy-mm-dd hh:mm');
}

function removeLogRecord_(practitionerNo, practiceNo, studentName, taskName, onlyStatus) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const logs = SpreadsheetApp.getActive().getSheetByName(DM.logs);
    const lastRow = findLastLogRow_(logs);
    if (lastRow < DM.logFirstRow) return { ok: false };
    const data = logs.getRange(DM.logFirstRow, 1, lastRow - DM.logFirstRow + 1, 9).getDisplayValues();
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      const isPenalty = String(row[5]).trim() === '' && String(row[6]).trim() === '' && Number(row[7]) === -2;
      const statusMatches = !onlyStatus || (onlyStatus === 'Штраф -2' ? isPenalty : String(row[8]).trim() === onlyStatus);
      const same = String(row[0]).trim() === studentName &&
        String(row[1]).trim() === DM.practitioners[practitionerNo - 1] &&
        Number(row[2]) === practiceNo && tasksEqual_(row[3], taskName) &&
        statusMatches;
      if (same) logs.getRange(DM.logFirstRow + i, 1, 1, 11).clearContent();
    }
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

function handleПродEdit_(commonRow) {
  const common = SpreadsheetApp.getActive().getSheetByName(DM.common);
  const status = String(common.getRange(commonRow, 4).getDisplayValue()).trim();
  if (status !== 'Нет') return;
  const studentName = String(common.getRange(commonRow, 1).getDisplayValue()).trim();
  const practitioner = String(common.getRange(commonRow, 2).getDisplayValue()).trim();
  if (!studentName || !practitioner || hasDepartureLog_(studentName, practitioner)) return;

  const logs = SpreadsheetApp.getActive().getSheetByName(DM.logs);
  const newRow = Math.max(findLastLogRow_(logs) + 1, DM.logFirstRow);
  logs.getRange(newRow, 1, 1, 11).setValues([[
    studentName, practitioner, '', '', '', '', '', '', 'ушел на базу', '', new Date()
  ]]);
  logs.getRange(newRow, 11).setNumberFormat('yyyy-mm-dd hh:mm');
}

function handlePractitionerEdit_(commonRow, event) {
  if (!event || !event.range || event.range.getNumRows() !== 1 || event.range.getNumColumns() !== 1) return;
  const oldPractitioner = String(event.oldValue || '').trim();
  const common = SpreadsheetApp.getActive().getSheetByName(DM.common);
  const newPractitioner = String(common.getRange(commonRow, 2).getDisplayValue()).trim();
  const studentName = String(common.getRange(commonRow, 1).getDisplayValue()).trim();
  if (!studentName || !oldPractitioner || !newPractitioner || oldPractitioner === newPractitioner) return;

  const logs = SpreadsheetApp.getActive().getSheetByName(DM.logs);
  const newRow = Math.max(findLastLogRow_(logs) + 1, DM.logFirstRow);
  logs.getRange(newRow, 1, 1, 11).setValues([[
    studentName, newPractitioner, '', '', '', '', '', '',
    `переведен из ${oldPractitioner} в ${newPractitioner}`, '', new Date()
  ]]);
  logs.getRange(newRow, 11).setNumberFormat('yyyy-mm-dd hh:mm');
}

function hasDepartureLog_(studentName, practitioner) {
  const logs = SpreadsheetApp.getActive().getSheetByName(DM.logs);
  const lastRow = findLastLogRow_(logs);
  if (lastRow < DM.logFirstRow) return false;
  return logs.getRange(DM.logFirstRow, 1, lastRow - DM.logFirstRow + 1, 9).getDisplayValues()
    .some(row => String(row[0]).trim() === studentName &&
      String(row[1]).trim() === practitioner &&
      ['ушел на базу', 'Ушёл на базу'].indexOf(String(row[8]).trim()) !== -1);
}

function refreshTotals() {
  const ss = SpreadsheetApp.getActive();
  const common = ss.getSheetByName(DM.common);
  for (let row = DM.commonFirstRow; row <= DM.commonLastRow; row++) {
    common.getRange(row, 5).setFormula(
      `=IF(A${row}="","",SUMIF('${DM.logs}'!$A$${DM.logFirstRow}:$A$${DM.logMaxRow},A${row},'${DM.logs}'!$H$${DM.logFirstRow}:$H$${DM.logMaxRow}))`
    );
    common.getRange(row, 6).setFormula(
      `=IF(A${row}="","",SUMIF('CW'!$A:$A,A${row},'CW'!$D:$D))`
    );
    common.getRange(row, 7).setFormula(
      `=IF(A${row}="","",IF(AND(E${row}>=55,F${row}>=24),"S",IF(AND(E${row}>=35,F${row}>=14),"A",IF(E${row}>=17,"B","Допса"))))`
    );
  }
  SpreadsheetApp.flush();
  refreshRanking();
}

function syncActiveStudents() {
  syncVisibility_();
}

function syncGroupAssignments() {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActive();
    const common = ss.getSheetByName(DM.common);
    const rosterRows = common.getRange(DM.commonFirstRow, 1, 90, 4).getDisplayValues();
    const groups = [[], [], []];
    rosterRows.forEach(row => {
      const practitionerNo = DM.practitioners.indexOf(String(row[1]).trim());
      const name = String(row[0]).trim();
      if (practitionerNo >= 0 && name) groups[practitionerNo].push({ name, tg: row[2] });
    });

    const snapshots = collectPlusData_();
    DM.plusSheets.forEach((sheetName, practitionerIndex) => {
      const sheet = ss.getSheetByName(sheetName);
      for (let practiceNo = 1; practiceNo <= DM.practiceCount; practiceNo++) {
        const blockStart = 4 + (practiceNo - 1) * DM.blockSize;
        const plusFirst = blockStart + 2;
        sheet.getRange(plusFirst, 1, DM.studentsPerGroup, 2).clearContent();
        sheet.getRange(plusFirst, DM.taskFirstCol, DM.studentsPerGroup, 30).clearContent();

        groups[practitionerIndex].slice(0, DM.studentsPerGroup).forEach((student, offset) => {
          const plusRow = plusFirst + offset;
          const saved = snapshots[student.name] || {};
          sheet.getRange(plusRow, 1, 1, 2).setValues([[student.name, student.tg]]);
          if (saved.plus && saved.plus[practiceNo]) sheet.getRange(plusRow, DM.taskFirstCol, 1, 30).setValues([saved.plus[practiceNo]]);
        });
      }
    });
    syncVisibility_();
    refreshTotals();
  } finally {
    lock.releaseLock();
  }
}

function collectPlusData_() {
  const ss = SpreadsheetApp.getActive();
  const snapshots = {};
  DM.plusSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    for (let practiceNo = 1; practiceNo <= DM.practiceCount; practiceNo++) {
      const blockStart = 4 + (practiceNo - 1) * DM.blockSize;
      const plusFirst = blockStart + 2;
      const plusRows = sheet.getRange(plusFirst, 1, DM.studentsPerGroup, 35).getValues();
      plusRows.forEach(row => {
        const name = String(row[0] == null ? '' : row[0]).trim();
        if (!name) return;
        snapshots[name] = snapshots[name] || { plus: {} };
        snapshots[name].plus[practiceNo] = row.slice(5, 35);
      });
    }
  });
  return snapshots;
}

function syncVisibility_() {
  const ss = SpreadsheetApp.getActive();
  const commonRows = ss.getSheetByName(DM.common).getRange(DM.commonFirstRow, 1, 90, 4).getDisplayValues();
  const activeByName = {};
  commonRows.forEach(row => { if (row[0]) activeByName[String(row[0]).trim()] = isActiveValue_(row[3]); });

  DM.plusSheets.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    for (let practiceNo = 1; practiceNo <= DM.practiceCount; practiceNo++) {
      const blockStart = 4 + (practiceNo - 1) * DM.blockSize;
      const plusFirst = blockStart + 2;
      const names = sheet.getRange(plusFirst, 1, DM.studentsPerGroup, 1).getDisplayValues();
      names.forEach((entry, offset) => {
        const active = entry[0] === '' || activeByName[String(entry[0]).trim()] !== false;
        if (active) {
          sheet.showRows(plusFirst + offset);
        } else {
          sheet.hideRows(plusFirst + offset);
        }
      });
    }
  });
}

function refreshRanking() {
  const ss = SpreadsheetApp.getActive();
  const common = ss.getSheetByName(DM.common);
  const ranking = ss.getSheetByName(DM.ranking);
  if (!ranking) return;
  const rows = common.getRange(DM.commonFirstRow, 1, 90, 7).getDisplayValues()
    .filter(row => String(row[0]).trim());
  rows.sort((a, b) => {
    const d = toNumber_(b[4]) - toNumber_(a[4]);
    if (d) return d;
    const cw = toNumber_(b[5]) - toNumber_(a[5]);
    if (cw) return cw;
    return String(a[0]).localeCompare(String(b[0]));
  });
  ranking.getRange(4, 1, 90, 7).clearContent();
  if (rows.length) ranking.getRange(4, 1, rows.length, 7).setValues(rows);
}

function setupValidations_() {
  const ss = SpreadsheetApp.getActive();
  const list = values => SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true).setAllowInvalid(false).build();

  DM.plusSheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    for (let p = 1; p <= DM.practiceCount; p++) {
      const blockStart = 4 + (p - 1) * DM.blockSize;
      const plusFirst = blockStart + 2;
      sheet.getRange(plusFirst, DM.taskFirstCol, DM.studentsPerGroup, 30)
        .setNumberFormat('@').setDataValidation(list(['+', '1', '-']));
    }
  });
  ss.getSheetByName(DM.common).getRange('B4:B93')
    .setDataValidation(list(DM.practitioners));
  ss.getSheetByName(DM.common).getRange('D4:D93')
    .setDataValidation(list(['Да', 'Нет']));
}

function applyConditionalFormatting_() {
  const ss = SpreadsheetApp.getActive();
  DM.plusSheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    let rules = sheet.getConditionalFormatRules().filter(rule => !rule.getRanges().some(r => r.getColumn() >= DM.taskFirstCol && r.getColumn() <= DM.taskLastCol));
    for (let p = 1; p <= DM.practiceCount; p++) {
      const blockStart = 4 + (p - 1) * DM.blockSize;
      const plusRange = sheet.getRange(blockStart + 2, DM.taskFirstCol, DM.studentsPerGroup, 30);
      rules.push(SpreadsheetApp.newConditionalFormatRule().setRanges([plusRange]).whenTextEqualTo('+').setBackground('#C6E0B4').setFontColor('#000000').build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().setRanges([plusRange]).whenTextEqualTo('1').setBackground('#548235').setFontColor('#FFFFFF').build());
      rules.push(SpreadsheetApp.newConditionalFormatRule().setRanges([plusRange]).whenTextEqualTo('-').setBackground('#C00000').setFontColor('#FFFFFF').build());
    }
    sheet.setConditionalFormatRules(rules);
  });
}

function findStudentRow_(sheet, studentName, firstRow, lastRow) {
  const values = sheet.getRange(firstRow, 1, lastRow - firstRow + 1, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) if (String(values[i][0]).trim() === studentName) return firstRow + i;
  return 0;
}

function findLastLogRow_(sheet) {
  const values = sheet.getRange(DM.logFirstRow, 1, DM.logMaxRow - DM.logFirstRow + 1, 1).getDisplayValues();
  let last = DM.logFirstRow - 1;
  values.forEach((row, i) => { if (String(row[0]).trim()) last = DM.logFirstRow + i; });
  return last;
}

function isStudentActive_(practitionerNo, studentName) {
  const common = SpreadsheetApp.getActive().getSheetByName(DM.common);
  const values = common.getRange(DM.commonFirstRow, 1, 90, 4).getDisplayValues();
  return values.some(row => DM.practitioners[practitionerNo - 1] === String(row[1]).trim() &&
    String(row[0]).trim() === studentName && isActiveValue_(row[3]));
}

function isActiveValue_(value) {
  const text = String(value == null ? '' : value).trim().toLowerCase();
  return !text || (text !== 'нет' && text !== 'no' && text !== 'n');
}

function isPractitionerEditor_() {
  if (!DM.practitionerEmails.length) return true;
  const email = Session.getActiveUser().getEmail();
  return !!email && DM.practitionerEmails.indexOf(email.toLowerCase()) !== -1;
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
  const number = Number(String(value == null ? '' : value).replace(',', '.'));
  return isNaN(number) ? 0 : number;
}
