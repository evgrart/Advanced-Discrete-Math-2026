const DM = {
  common: 'Общий',
  logs: 'Логи',
  plusSheets: ['Плюсы_Артём', 'Плюсы_Рами', 'Плюсы_Немат'],
  practiceCount: 15,
  blockSize: 68,
  taskFirstCol: 6,
  taskLastCol: 35,
  studentsPerGroup: 30,
  logFirstRow: 7,
  logMaxRow: 3006
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('DM система')
    .addItem('Настроить таблицу', 'installDmSystem')
    .addItem('Синхронизировать активность', 'syncActiveStudents')
    .addItem('Пересчитать D', 'refreshTotals')
    .addToUi();
}

function installDmSystem() {
  setupValidations_();
  refreshTotals();
  syncActiveStudents();
  SpreadsheetApp.getActive().toast('DM-система установлена. Ставьте 1 вместо + в верхних плюсовиках.');
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

      if (sheet.getName() === DM.common && col === 4 && row >= 4 && row <= 93) {
        syncOneStudent(row);
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
    const plusFirst = 6 + (practiceNo - 1) * DM.blockSize;
    const plusLast = plusFirst + DM.studentsPerGroup - 1;
    if (row < plusFirst || row > plusLast) continue;

    const value = String(cell.getDisplayValue()).trim();
    const studentName = String(sheet.getRange(row, 1).getDisplayValue()).trim();
    const headerRow = 5 + (practiceNo - 1) * DM.blockSize;
    const taskName = String(sheet.getRange(headerRow, col).getDisplayValue()).trim();
    if (!studentName || !taskName) return;

    if (value === '1') {
      const result = registerUpperOne_(practitionerNo, practiceNo, studentName, taskName);
      cell.setNote(result.ok ? 'Выход добавлен в «Логи».' : result.message);
    } else if (event && String(event.oldValue || '').trim() === '1' && value !== '1') {
      const result = removeLogRecord_(practitionerNo, practiceNo, studentName, taskName);
      cell.setNote(result.ok ? 'Выход удалён из «Логи».' : result.message);
    }
    return;
  }
}

function registerUpperOne_(practitionerNo, practiceNo, studentName, taskName) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.getActive();
    if (!isStudentActive_(practitionerNo, studentName)) {
      return { ok: false, message: 'Студент неактивен на листе «Общий».' };
    }

    const plusSheet = ss.getSheetByName(DM.plusSheets[practitionerNo - 1]);
    const result = appendLogRow_(ss.getSheetByName(DM.logs), plusSheet,
      practitionerNo, practiceNo, studentName, taskName);
    SpreadsheetApp.flush();
    return result;
  } finally {
    lock.releaseLock();
  }
}

function appendLogRow_(logs, plusSheet, practitionerNo, practiceNo, studentName, taskName) {
  const lastRow = findLastLogRow_(logs);
  const data = lastRow >= DM.logFirstRow
    ? logs.getRange(DM.logFirstRow, 1, lastRow - DM.logFirstRow + 1, 4).getDisplayValues()
    : [];
  let studentExit = 0;
  let pairCount = 0;
  let duplicateCount = 0;

  data.forEach(row => {
    if (String(row[2]).trim() !== studentName) return;
    studentExit++;
    if (Number(row[0]) === practiceNo) pairCount++;
    if (Number(row[0]) === practiceNo && tasksEqual_(row[3], taskName)) duplicateCount++;
  });
  studentExit++;

  if (pairCount >= 2) return { ok: false, message: 'У студента уже два выхода за эту практику.' };
  if (duplicateCount > 0) return { ok: false, message: 'Эта задача уже есть в «Логи».' };

  const plusHeader = 5 + (practiceNo - 1) * DM.blockSize;
  const studentRow = findStudentRow_(plusSheet, studentName, plusHeader + 1, plusHeader + 30);
  if (!studentRow) return { ok: false, message: 'Студент не найден на плюсовике.' };
  const coefficient = Number(plusSheet.getRange(studentRow, 4).getValue()) || 0.5;
  const base = ladderPoints_(studentExit);
  const newRow = Math.max(findLastLogRow_(logs) + 1, DM.logFirstRow);

  logs.getRange(newRow, 1, 1, 9).setValues([[
    practiceNo, practitionerNo, studentName, taskName, studentExit,
    base, coefficient, '', studentExit > 15 ? 'Extra output' : 'Assigned'
  ]]);
  logs.getRange(newRow, 8).setFormula(`=F${newRow}*G${newRow}+IF(J${newRow}="",0,J${newRow})`);
  return { ok: true, total: base * coefficient };
}

function removeLogRecord_(practitionerNo, practiceNo, studentName, taskName) {
  const lock = LockService.getDocumentLock();
  lock.waitLock(30000);
  try {
    const logs = SpreadsheetApp.getActive().getSheetByName(DM.logs);
    const lastRow = findLastLogRow_(logs);
    if (lastRow < DM.logFirstRow) return { ok: false, message: 'Запись не найдена.' };
    const data = logs.getRange(DM.logFirstRow, 1, lastRow - DM.logFirstRow + 1, 4).getDisplayValues();
    for (let i = 0; i < data.length; i++) {
      const row = data[i];
      if (Number(row[0]) === practiceNo && Number(row[1]) === practitionerNo &&
          String(row[2]).trim() === studentName && tasksEqual_(row[3], taskName)) {
        logs.getRange(DM.logFirstRow + i, 1, 1, 10).clearContent();
        return { ok: true };
      }
    }
    return { ok: false, message: 'Запись не найдена.' };
  } finally {
    lock.releaseLock();
  }
}

function refreshTotals() {
  const ss = SpreadsheetApp.getActive();
  const common = ss.getSheetByName(DM.common);
  for (let row = 4; row <= 93; row++) {
    common.getRange(row, 5).setFormula(
      `=IF(B${row}="","",SUMIF('Логи'!$C$7:$C$3006,B${row},'Логи'!$H$7:$H$3006))`
    );
  }
  SpreadsheetApp.flush();
}

function syncActiveStudents() {
  for (let row = 4; row <= 93; row++) syncOneStudent(row);
}

function syncOneStudent(commonRow) {
  if (commonRow < 4 || commonRow > 93) return;
  const ss = SpreadsheetApp.getActive();
  const common = ss.getSheetByName(DM.common);
  const practitionerNo = Math.floor((commonRow - 4) / DM.studentsPerGroup) + 1;
  const offset = (commonRow - 4) % DM.studentsPerGroup;
  const active = isActiveValue_(common.getRange(commonRow, 4).getValue());
  const sheet = ss.getSheetByName(DM.plusSheets[practitionerNo - 1]);

  for (let practiceNo = 1; practiceNo <= DM.practiceCount; practiceNo++) {
    const plusRow = 6 + (practiceNo - 1) * DM.blockSize + offset;
    const resultRow = 39 + (practiceNo - 1) * DM.blockSize + offset;
    if (active) {
      sheet.showRows(plusRow);
      sheet.showRows(resultRow);
    } else {
      sheet.hideRows(plusRow);
      sheet.hideRows(resultRow);
    }
  }
}

function setupValidations_() {
  const ss = SpreadsheetApp.getActive();
  const list = values => SpreadsheetApp.newDataValidation()
    .requireValueInList(values, true).setAllowInvalid(false).build();

  DM.plusSheets.forEach(name => {
    const sheet = ss.getSheetByName(name);
    for (let p = 1; p <= DM.practiceCount; p++) {
      const plusFirst = 6 + (p - 1) * DM.blockSize;
      const resultFirst = 39 + (p - 1) * DM.blockSize;
      sheet.getRange(plusFirst, DM.taskFirstCol, DM.studentsPerGroup, 30)
        .setDataValidation(list(['+', '1']));
      sheet.getRange(resultFirst, DM.taskFirstCol, DM.studentsPerGroup, 30)
        .setDataValidation(list(['1', '-']));
    }
  });
  ss.getSheetByName(DM.common).getRange('D4:D93')
    .setDataValidation(list(['Да', 'Нет']));
}

function findStudentRow_(sheet, studentName, firstRow, lastRow) {
  const values = sheet.getRange(firstRow, 1, lastRow - firstRow + 1, 1).getDisplayValues();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0]).trim() === studentName) return firstRow + i;
  }
  return 0;
}

function findLastLogRow_(sheet) {
  const values = sheet.getRange(DM.logFirstRow, 3,
    DM.logMaxRow - DM.logFirstRow + 1, 1).getDisplayValues();
  let last = DM.logFirstRow - 1;
  values.forEach((row, i) => { if (String(row[0]).trim()) last = DM.logFirstRow + i; });
  return last;
}

function isStudentActive_(practitionerNo, studentName) {
  const common = SpreadsheetApp.getActive().getSheetByName(DM.common);
  const first = 4 + (practitionerNo - 1) * DM.studentsPerGroup;
  const values = common.getRange(first, 2, DM.studentsPerGroup, 3).getDisplayValues();
  for (const row of values) {
    if (String(row[0]).trim() === studentName) return isActiveValue_(row[2]);
  }
  return false;
}

function isActiveValue_(value) {
  const text = String(value == null ? '' : value).trim().toLowerCase();
  return !text || (text !== 'нет' && text !== 'no' && text !== 'n');
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
