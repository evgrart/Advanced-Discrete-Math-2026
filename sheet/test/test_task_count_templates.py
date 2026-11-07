import importlib.util
from pathlib import Path
import tempfile
import unittest


ROOT = Path(__file__).resolve().parents[1]
spec = importlib.util.spec_from_file_location("builder", ROOT / "build_four_workbooks.py")
builder = importlib.util.module_from_spec(spec)
spec.loader.exec_module(builder)

from openpyxl import load_workbook
from openpyxl import Workbook
from openpyxl.utils import get_column_letter


class TaskCountTemplatesTest(unittest.TestCase):
    def test_control_work_thresholds_are_read_from_shared_config(self):
        self.assertEqual(set(builder.load_cw_settings()), {1, 2})
        with tempfile.TemporaryDirectory() as directory:
            config = Path(directory) / "Code.gs"
            config.write_text("const DM = {cwRequiredSolvedTasks: {1: 3, 2: 5}};", encoding="utf8")
            self.assertEqual(builder.load_cw_settings(config), {1: 3, 2: 5})
            for mapping in ['1: 4', '1: 4, 2: 7', '1: 0, 2: 4', '1: 3.5, 2: 4',
                            '1: 4, 2: "4"', '1:4, 1:5, 2:4', '1:4, 2:4, 3:4']:
                config.write_text("const DM = {cwRequiredSolvedTasks: {" + mapping + "}};", encoding="utf8")
                with self.subTest(mapping=mapping), self.assertRaises(ValueError):
                    builder.load_cw_settings(config)

    def test_central_templates_have_six_scored_tasks_and_independent_pass_thresholds(self):
        wb = Workbook()
        wb.remove(wb.active)
        thresholds = {1: 3, 2: 5}
        common = builder.make_common(wb, thresholds)
        cw = builder.make_cw(wb, thresholds)
        for kr, header, first, last in builder.CW_BLOCKS:
            self.assertEqual([cw.cell(header, col).value for col in range(5, 11)], list("123456"))
            self.assertIn(f"SUM(E{first}:J{first})", cw.cell(first, 4).value)
            self.assertIn(f'1/2^(C{last}-1)', cw.cell(last, 4).value)
            self.assertIn(f"$J${first}:$J${last}", common.cell(4, kr + 6).value)
            self.assertIn(f'>={thresholds[kr]},"зачет"', common.cell(4, kr + 6).value)
            rules = cw.conditional_formatting[f"A{first}:A{last}"]
            self.assertEqual(rules[0].formula, [f'AND($A{first}<>"",COUNTIF(E{first}:J{first},">=3")>={thresholds[kr]})'])
            self.assertTrue(any(validation.type == "decimal" and validation.formula1 == "0" and
                                validation.formula2 == "4" and f"J{first}" in validation.sqref
                                for validation in cw.data_validations.dataValidation))
            self.assertFalse(cw.cell(first, 10).protection.locked)
        self.assertIn('>=55', common['I4'].value)
        self.assertIn('>=24', common['I4'].value)
        self.assertIn('>=35', common['I4'].value)
        self.assertIn('>=14', common['I4'].value)
        self.assertIn('>=17', common['I4'].value)
        wb.close()

    def test_upgrading_old_central_workbook_preserves_scores_attempts_and_unrelated_cells(self):
        wb = Workbook()
        wb.remove(wb.active)
        builder.make_common(wb)
        cw = builder.make_cw(wb)
        cw.delete_cols(10)
        wb.create_sheet('Логи')['A4'] = 'Неприкосновенная запись'
        for first, attempt, scores in [(3, 2, [4, 3.5, 1, 0, 2]), (96, 4, [0, 1, 2, 3, 4])]:
            cw.cell(first, 1, 'Студент')
            cw.cell(first, 3, attempt)
            for col, value in enumerate(scores, 5):
                cw.cell(first, col, value)
        before = {(sheet.title, cell.coordinate): cell.value for sheet in wb for row in sheet for cell in row}
        builder.update_central_control_works(wb, {1: 3, 2: 5})
        changed = {('CW', f'D{row}') for _, _, first, last in builder.CW_BLOCKS for row in range(first, last + 1)}
        changed.update(('Общий', f'{col}{row}') for row in range(4, 94) for col in 'GH')
        for (sheet, coordinate), value in before.items():
            if (sheet, coordinate) not in changed:
                self.assertEqual(wb[sheet][coordinate].value, value, (sheet, coordinate))
        cw['J3'] = 4
        cw['J96'] = 3.5
        count = len(cw.data_validations.dataValidation)
        builder.update_central_control_works(wb, {1: 3, 2: 5})
        self.assertEqual(len(cw.data_validations.dataValidation), count)
        self.assertEqual((cw['J3'].value, cw['J96'].value), (4, 3.5))
        with tempfile.TemporaryDirectory() as directory:
            path = Path(directory) / 'central.xlsx'
            wb.save(path)
            actual = load_workbook(path)
            self.assertEqual((actual['CW']['J3'].value, actual['CW']['J96'].value), (4, 3.5))
            self.assertEqual(actual['CW']['C3'].value, 2)
            self.assertEqual(actual['CW']['C96'].value, 4)
            self.assertIn('SUM(E96:J96)', actual['CW']['D96'].value)
            actual.close()
        wb.close()

    def test_reads_shared_code_config_and_rejects_malformed_values(self):
        with tempfile.TemporaryDirectory() as directory:
            config = Path(directory) / "Code.gs"
            config.write_text("const DM = {defaultTaskCount: 30, practiceTaskCounts: {1:30, 2:25, 3:40}};", encoding="utf8")
            counts = builder.load_task_counts(config)
            self.assertEqual([counts[i] for i in range(1, 5)], [30, 25, 40, 30])
            for mapping in ['2: 0', '2: -1', '2: 2.5', '2: "25"', '2: 201', '16: 25', '"02":25', '2:25,2:40']:
                config.write_text("const DM = {defaultTaskCount: 30, practiceTaskCounts: {" + mapping + "}};", encoding="utf8")
                with self.subTest(mapping=mapping), self.assertRaises(ValueError):
                    builder.load_task_counts(config)
        self.assertEqual(len(builder.load_task_counts()), 15)

    def test_builds_three_workbooks_with_30_25_40_and_shared_formulas(self):
        counts = {practice: 30 for practice in range(1, 16)}
        counts.update({2: 25, 3: 40})
        with tempfile.TemporaryDirectory() as directory:
            for index, name in enumerate(builder.PRACTITIONERS, 1):
                path = Path(directory) / f"group-{index}.xlsx"
                builder.make_plus(path, name, index, counts)
                workbook = load_workbook(path)
                first = 5 if name == 'Рами' else 6
                solved = 2 if name == 'Рами' else 3
                for practice, count in counts.items():
                    sheet = workbook[f"Практика {practice}"]
                    last = first + count - 1
                    self.assertEqual(sheet.max_column, last)
                    self.assertEqual([sheet.cell(3, c).value for c in range(first, last + 1)], [str(i) for i in range(1, count + 1)])
                    self.assertIn(f"{get_column_letter(first)}4:{get_column_letter(last)}4", sheet.cell(4, solved).value)
                    self.assertIn(f"{get_column_letter(first)}$3:{get_column_letter(last)}$3", sheet.cell(4, solved + 1).value)
                    rules = [rule for group in sheet.conditional_formatting for rule in sheet.conditional_formatting[group]]
                    self.assertTrue(any(rule.formula == [f"2*{count}/3"] for rule in rules))
                    self.assertTrue(any(rule.formula == [f"{count}/2", f"2*{count}/3"] for rule in rules))
                    self.assertFalse(sheet.protection.sheet)
                    self.assertIsNone(sheet.cell(4, last).value)
                self.assertIn('Логи', workbook.sheetnames)
                workbook.close()


if __name__ == '__main__':
    unittest.main()
