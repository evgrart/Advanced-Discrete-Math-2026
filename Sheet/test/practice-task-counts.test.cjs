const assert = require('node:assert/strict');
const fs = require('node:fs');
const path = require('node:path');
const vm = require('node:vm');
const test = require('node:test');

class Range {
  constructor(sheet, row, column, height = 1, width = 1) {
    assert.ok(row > 0 && column > 0 && height > 0 && width > 0);
    assert.ok(column + width - 1 <= sheet.maxColumns, 'read beyond grid');
    Object.assign(this, { sheet, row, column, height, width });
  }
  getValues() {
    return Array.from({ length: this.height }, (_, y) =>
      Array.from({ length: this.width }, (_, x) => this.sheet.data[this.row + y - 1]?.[this.column + x - 1] ?? ''));
  }
  getDisplayValues() { return this.getValues().map(row => row.map(String)); }
  getFormulas() {
    return Array.from({ length: this.height }, (_, y) =>
      Array.from({ length: this.width }, (_, x) => this.sheet.formulas.get(`${this.row + y}:${this.column + x}`) || ''));
  }
  setValues(values) {
    assert.equal(values.length, this.height);
    values.forEach((row, y) => {
      assert.equal(row.length, this.width);
      row.forEach((value, x) => {
        this.sheet.put(this.row + y, this.column + x, value);
        this.sheet.formulas.delete(`${this.row + y}:${this.column + x}`);
      });
    });
    this.sheet.writes++;
    return this;
  }
  setFormulas(values) {
    assert.equal(values.length, this.height);
    values.forEach((row, y) => {
      assert.equal(row.length, this.width);
      row.forEach((value, x) => this.sheet.formulas.set(`${this.row + y}:${this.column + x}`, value));
    });
    this.sheet.writes++;
    return this;
  }
  clearContent() { return this.setValues(this.getValues().map(row => row.map(() => ''))); }
  getColumn() { return this.column; }
  getLastColumn() { return this.column + this.width - 1; }
  getNumColumns() { return this.width; }
  getRow() { return this.row; }
  getNumRows() { return this.height; }
  getSheet() { return this.sheet; }
  getDisplayValue() { return this.getDisplayValues()[0][0]; }
  clearNote() { return this; }
  setNumberFormat() { return this; }
  setBackground() { return this; }
  setFontColor() { return this; }
  setFontWeight() { return this; }
  copyFormatToRange(sheet, from, to) { sheet.copiedFormat = [from, to]; return this; }
}

class Sheet {
  constructor(name, maxColumns = 35) {
    Object.assign(this, { name, maxColumns, data: [], formulas: new Map(), rules: [], hidden: new Set(), writes: 0 });
  }
  put(row, col, value) {
    if (!this.data[row - 1]) this.data[row - 1] = [];
    this.data[row - 1][col - 1] = value;
  }
  getRange(...args) { return new Range(this, ...args); }
  getName() { return this.name; }
  getLastRow() { return Math.max(5, this.data.length); }
  getLastColumn() {
    return Math.max(1, ...this.data.filter(Boolean).map(row => row.reduce((last, value, i) => value !== '' && value != null ? i + 1 : last, 0)));
  }
  getMaxColumns() { return this.maxColumns; }
  insertColumnsAfter(after, number) { assert.equal(after, this.maxColumns); this.maxColumns += number; this.writes++; }
  getColumnWidth() { return 80; }
  setColumnWidths() {}
  hideColumns(first, count) { for (let i = first; i < first + count; i++) this.hidden.add(i); }
  showColumns(first, count) { for (let i = first; i < first + count; i++) this.hidden.delete(i); }
  getConditionalFormatRules() { return this.rules; }
  setConditionalFormatRules(rules) { this.rules = rules; }
  showRows() {}
  hideRows() {}
}

function ruleBuilder() {
  return {
    whenFormulaSatisfied(formula) { this.formula = formula; return this; },
    setRanges(ranges) { this.ranges = ranges; return this; },
    setBackground() { return this; }, setFontColor() { return this; },
    build() {
      const { ranges, formula } = this;
      return { formula, getRanges: () => ranges, copy: () => ruleBuilder().whenFormulaSatisfied(formula).setRanges(ranges) };
    }
  };
}

function setup() {
  const context = vm.createContext({
    SpreadsheetApp: { newConditionalFormatRule: ruleBuilder, flush() {} },
    LockService: { getScriptLock: () => ({ waitLock() {}, releaseLock() {} }) }
  });
  vm.runInContext(fs.readFileSync(path.join(__dirname, '../Code.gs'), 'utf8'), context);
  const dm = vm.runInContext('DM', context);
  Object.assign(dm, { practiceCount: 3, studentsPerGroup: 2, commonLastRow: 5, logMaxRow: 20 });
  const plusFiles = {};
  for (const practitioner of dm.practitioners) {
    const first = practitioner === 'Рами' ? 5 : 6;
    const sheets = {};
    for (let practice = 1; practice <= dm.practiceCount; practice++) {
      const sheet = new Sheet(`Практика ${practice}`, first + 29);
      for (let n = 1; n <= 30; n++) sheet.put(3, first + n - 1, n);
      sheet.put(4, 1, 'Студент'); sheet.put(5, 1, 'Второй');
      sheets[sheet.name] = sheet;
    }
    plusFiles[practitioner] = { getSheetByName: name => sheets[name] || null };
  }
  const logs = new Sheet('Логи', 11);
  const central = { getSheetByName: name => name === 'Логи' ? logs : null };
  return { context, dm, plusFiles, central, logs,
    sheet: (name, practice) => plusFiles[name].getSheetByName(`Практика ${practice}`) };
}

test('config applies shared counts, default, different Rami offset, and validates inputs', () => {
  const { context: c, dm } = setup();
  dm.practiceTaskCounts = { 1: 30, 2: 25, 3: 40 };
  c.validatePracticeTaskCounts_();
  for (const practitioner of dm.practitioners) {
    for (const [practice, count] of Object.entries(dm.practiceTaskCounts)) {
      const layout = c.plusLayout_(practitioner, Number(practice));
      assert.equal(layout.taskLastCol - layout.taskFirstCol + 1, count);
      assert.equal(layout.taskFirstCol, practitioner === 'Рами' ? 5 : 6);
    }
  }
  for (const invalid of [{ 2: 0 }, { 2: -1 }, { 2: 2.5 }, { 2: '25' }, { 2: 201 }, { 4: 25 }, { '02': 25 }, []]) {
    dm.practiceTaskCounts = invalid;
    assert.throws(() => c.validatePracticeTaskCounts_());
  }
  dm.practiceTaskCounts = {};
  assert.equal(c.practiceTaskCount_(2), 30);
});

test('shrinks and expands all practitioners, preserves marks, adjusts formulas/colors, is idempotent', () => {
  const { context: c, dm, plusFiles, central, sheet } = setup();
  dm.practiceTaskCounts = { 2: 25, 3: 40 };
  for (const name of dm.practitioners) {
    const layout = c.plusLayout_(name, 2);
    const s = sheet(name, 2);
    s.put(4, layout.taskFirstCol, '+'); s.put(5, layout.taskFirstCol + 24, 1);
    s.rules = [ruleBuilder().setRanges([s.getRange(4, 1, 2, 1)]).whenFormulaSatisfied('=$A4<>""').build()];
  }
  c.syncPracticeTaskLayouts_(central, plusFiles);
  for (const name of dm.practitioners) {
    const l = c.plusLayout_(name, 2); const s = sheet(name, 2);
    assert.equal(c.practiceTaskHeaders_(s, l).length, 25);
    assert.equal(s.getRange(4, l.taskFirstCol).getValues()[0][0], '+');
    assert.equal(s.getRange(5, l.taskLastCol).getValues()[0][0], 1);
    assert.ok(s.hidden.has(l.taskLastCol + 1));
    const formulas = s.getRange(4, l.solvedCol, 1, 2).getFormulas()[0];
    assert.ok(formulas[0].includes(`${c.columnLabel_(l.taskLastCol)}4`));
    assert.ok(formulas[1].includes('2*25/3'));
    assert.ok(formulas[1].includes('>=25/2'));
    assert.ok(s.rules.some(rule => rule.formula === '=$A4<>""'));
    assert.ok(s.rules.some(rule => rule.formula.includes('2*25/3')));
    const expanded = sheet(name, 3), el = c.plusLayout_(name, 3);
    assert.equal(expanded.getMaxColumns(), el.taskLastCol);
    assert.equal(expanded.getRange(3, el.taskLastCol).getDisplayValue(), '40');
    assert.equal(expanded.getRange(4, el.taskLastCol).getDisplayValue(), '');
    assert.equal(sheet(name, 1).writes, 0);
  }
  const writes = dm.practitioners.map(n => sheet(n, 2).writes + sheet(n, 3).writes);
  c.syncPracticeTaskLayouts_(central, plusFiles);
  assert.deepEqual(dm.practitioners.map(n => sheet(n, 2).writes + sheet(n, 3).writes), writes);
});

for (const marker of ['+', 1, '-', 'note']) {
  test(`refuses shrink with ${marker} in last practitioner before any writes`, () => {
    const { context: c, dm, plusFiles, central, sheet } = setup();
    dm.practiceTaskCounts = { 2: 25 };
    sheet('Немат', 2).put(4, 35, marker);
    assert.throws(() => c.syncPracticeTaskLayouts_(central, plusFiles), /есть данные/);
    for (const name of dm.practitioners) assert.equal(sheet(name, 2).writes, 0);
  });
}

test('refuses shrink with blank-result formula or historical central log', () => {
  const { context: c, dm, plusFiles, central, sheet, logs } = setup();
  dm.practiceTaskCounts = { 2: 25 };
  sheet('Рами', 2).formulas.set('4:34', '=IF(TRUE,"",1)');
  assert.throws(() => c.syncPracticeTaskLayouts_(central, plusFiles), /есть данные/);
  sheet('Рами', 2).formulas.clear();
  [ 'Студент', 'Немат', 2, 30 ].forEach((v, i) => logs.put(4, i + 1, v));
  assert.throws(() => c.syncPracticeTaskLayouts_(central, plusFiles), /Логах/);
  for (const name of dm.practitioners) assert.equal(sheet(name, 2).writes, 0);
});

test('refuses to turn unlabelled data into new task marks', () => {
  const { context: c, dm, plusFiles, central, sheet } = setup();
  dm.practiceTaskCounts = { 3: 40 };
  const s = sheet('Немат', 3);
  s.maxColumns = 45;
  s.put(4, 36, '+');
  assert.throws(() => c.syncPracticeTaskLayouts_(central, plusFiles), /без заголовков/);
  for (const name of dm.practitioners) assert.equal(sheet(name, 3).writes, 0);
});

test('force reapplies formulas with unchanged counts; shrink/grow round trip keeps marks', () => {
  const { context: c, dm, plusFiles, central, sheet } = setup();
  dm.practiceTaskCounts = { 2: 25 };
  c.syncPracticeTaskLayouts_(central, plusFiles);
  for (const name of dm.practitioners) {
    const l = c.plusLayout_(name, 2); const s = sheet(name, 2);
    s.put(4, l.taskFirstCol, '+');
    s.formulas.set(`4:${l.solvedCol}`, '=0');
  }
  c.syncPracticeTaskLayouts_(central, plusFiles, true);
  for (const name of dm.practitioners) {
    const l = c.plusLayout_(name, 2);
    assert.ok(sheet(name, 2).formulas.get(`4:${l.solvedCol}`).includes('COUNTIF'));
  }
  dm.practiceTaskCounts = { 2: 30 };
  c.syncPracticeTaskLayouts_(central, plusFiles);
  for (const name of dm.practitioners) {
    const l = c.plusLayout_(name, 2); const s = sheet(name, 2);
    assert.equal(s.getRange(4, l.taskFirstCol).getDisplayValue(), '+');
    assert.equal(s.getRange(4, l.taskLastCol).getDisplayValue(), '');
    assert.ok(!s.hidden.has(l.taskLastCol));
    assert.equal(c.practiceTaskHeaders_(s, l).length, 30);
    assert.equal(s.rules.length, 7);
  }
});

test('pending counts reject edits instead of deleting or creating log entries', () => {
  const { context: c, dm, sheet } = setup();
  dm.practiceTaskCounts = { 2: 25 };
  const s = sheet('Артём', 2);
  const event = { range: s.getRange(4, 6), source: {} };
  assert.throws(() => c.handlePlusEdit_(event, 'Артём'), /Применить количество задач/);
  assert.equal(s.writes, 0);
});

test('onEdit logs task 40 after expansion for each practitioner', () => {
  const { context: c, dm, plusFiles, central, sheet } = setup();
  dm.practiceTaskCounts = { 3: 40 };
  c.syncPracticeTaskLayouts_(central, plusFiles);
  dm.centralFileId = 'central';
  c.SpreadsheetApp.openById = () => central;
  c.removeLogRecord_ = () => {}; c.refreshCentral_ = () => {};
  const appended = [];
  c.appendNormalLog_ = (...args) => { appended.push([args[4], args[5], args[7]]); return { ok: true }; };
  for (const name of dm.practitioners) {
    const s = sheet(name, 3), layout = c.plusLayout_(name, 3);
    s.put(4, layout.taskLastCol, '1');
    c.handlePlusEdit_({ range: s.getRange(4, layout.taskLastCol), source: plusFiles[name], oldValue: '+' }, name);
  }
  assert.deepEqual(appended, [['Артём', 3, '40'], ['Рами', 3, '40'], ['Немат', 3, '40']]);
});

test('group synchronization uses per-practice widths and preserves existing marks', () => {
  const { context: c, dm, plusFiles, sheet, logs } = setup();
  dm.practiceTaskCounts = { 2: 25, 3: 40 };
  const common = new Sheet('Общий', 14);
  dm.practitioners.forEach((name, i) => {
    [name + ' студент', name, '', 'Да'].forEach((v, j) => common.put(4 + i, 1 + j, v));
    for (let practice = 1; practice <= 3; practice++) {
      const s = sheet(name, practice), l = c.plusLayout_(name, practice);
      s.put(4, 1, name + ' студент'); s.put(5, 1, '');
      s.put(4, l.taskFirstCol, '+');
    }
  });
  let cwCalled = false, refreshed = false, reconciled = false;
  c.syncCwRoster_ = () => { cwCalled = true; };
  c.refreshCentral_ = () => { refreshed = true; };
  c.reconcilePlusResults_ = (_central, _files, states) => {
    reconciled = true;
    for (const name of dm.practitioners) {
      assert.equal(states[name][2].tasks[0].length, 25);
      assert.equal(states[name][3].tasks[0].length, 40);
      assert.equal(states[name][2].tasks[0][0], '+');
      assert.equal(states[name][3].tasks[0][39], '');
    }
  };
  c.syncGroups_({ getSheetByName: name => name === 'Общий' ? common : logs }, plusFiles);
  assert.ok(cwCalled && reconciled && refreshed);
  dm.practiceTaskCounts[3] = 35;
  sheet('Немат', 3).put(4, 45, '+');
  cwCalled = false; reconciled = false; refreshed = false;
  assert.throws(() => c.syncGroups_({ getSheetByName: name => name === 'Общий' ? common : logs }, plusFiles), /есть данные/);
  assert.ok(!cwCalled && !reconciled && !refreshed);
});

test('snapshot, reconciliation and notes include tasks beyond 30 and exclude removed columns', () => {
  const { context: c, dm, plusFiles, central, sheet } = setup();
  dm.practiceTaskCounts = { 2: 25, 3: 40 };
  c.syncPracticeTaskLayouts_(central, plusFiles);
  sheet('Артём', 3).put(4, 45, 1);
  sheet('Рами', 3).put(5, 43, '-');
  const snapshots = c.collectSnapshots_(plusFiles);
  assert.equal(snapshots.sheets['Рами'][2].tasks[0].length, 25);
  assert.equal(snapshots.sheets['Артём'][3].tasks[0].length, 40);
  const appended = [];
  c.appendNormalLog_ = (...args) => { appended.push(['normal', args[5], args[6], args[7]]); return { ok: true }; };
  c.appendPenaltyLog_ = (...args) => { appended.push(['penalty', args[2], args[3], args[4]]); };
  c.reconcilePlusResults_(central, plusFiles, snapshots.sheets);
  assert.deepEqual(appended, [['normal', 3, 'Студент', '40'], ['penalty', 3, 'Второй', '39']]);
  c.clearPlusNotes_(plusFiles);
});

test('all coefficient boundaries use the configured count, including 1, 25, 30 and 40', () => {
  const { context: c, dm } = setup();
  for (const count of [1, 25, 30, 40, 200]) {
    dm.practiceTaskCounts = { 2: count };
    const layout = c.plusLayout_('Рами', 2);
    const formula = c.plusCountFormulas_(layout, count)[0][1];
    const match = formula.match(/IF\(B4>2\*(\d+)\/3,1,IF\(B4>=(\d+)\/2,0\.8,0\.5\)\)/);
    assert.ok(match);
    assert.equal(Number(match[1]), count); assert.equal(Number(match[2]), count);
    for (let solved = 0; solved <= count; solved++) {
      const coefficient = solved > 2 * Number(match[1]) / 3 ? 1 : solved >= Number(match[2]) / 2 ? .8 : .5;
      assert.equal(coefficient, solved * 3 > count * 2 ? 1 : solved * 2 >= count ? .8 : .5);
    }
  }
});

test('D formulas still trim names after synchronization', () => {
  const { context: c, dm } = setup();
  const common = new Sheet('Общий', 14);
  c.openPlusFiles_ = () => ({}); c.refreshLogCoefficients_ = () => {};
  c.refreshRanking_ = () => {}; c.refreshPlusTotals_ = () => {};
  c.refreshCentral_({ getSheetByName: () => common });
  assert.equal(common.getRange(4, 5).getFormulas()[0][0], '=IF(A4="","",SUMIF(\'Логи\'!$A$4:$A$20,TRIM(A4),\'Логи\'!$H$4:$H$20))');
});

// Optional real bridge integration: pass the path of the separate private repository.
if (process.env.DM_BOT_BRIDGE_PATH) {
  for (const name of ['Артём', 'Рами', 'Немат']) {
    test(`real Apps Script bridge reads mixed lengths for ${name}`, () => {
      const { context: c, dm, plusFiles, central, sheet } = setup();
      dm.practiceTaskCounts = { 2: 25, 3: 40 };
      c.syncPracticeTaskLayouts_(central, plusFiles);
      const common = new Sheet('Общий', 14);
      ['ФИО', 'Практик', 'TG', 'Прод'].forEach((v, i) => common.put(3, i + 1, v));
      ['Студент', name, '@student', 'Да'].forEach((v, i) => common.put(4, i + 1, v));
      for (let practice = 1; practice <= 3; practice++) {
        const layout = c.plusLayout_(name, practice); const s = sheet(name, practice);
        if (layout.presenceCol) s.put(4, layout.presenceCol, 'Да');
        for (let col = layout.taskFirstCol; col <= layout.taskLastCol; col++) s.put(4, col, '+');
      }
      dm.centralFileId = 'central'; dm.practitioners.forEach(n => dm.plusFileIds[n] = n);
      c.SpreadsheetApp.openById = id => id === 'central' ? { getSheetByName: () => common } : plusFiles[id];
      vm.runInContext(fs.readFileSync(process.env.DM_BOT_BRIDGE_PATH, 'utf8'), c);
      for (const practice of [1, 2, 3]) {
        const result = c.bridge_readDistributionProblem_(dm.practitioners.indexOf(name) + 1, practice);
        assert.equal(result.tasks.length, practice === 1 ? 30 : practice === 2 ? 25 : 40);
        assert.equal(result.students[0].semesterSolved, 95);
        assert.equal(result.students[0].candidateTaskKeys.length, result.tasks.length);
        assert.equal(result.students[0].slots, 2);
      }
      dm.practiceTaskCounts[2] = 24;
      assert.throws(() => c.bridge_readDistributionProblem_(dm.practitioners.indexOf(name) + 1, 2), /Применить количество задач/);
    });
  }
}
