from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent / ".deps"))

from openpyxl import Workbook
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Protection, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation


ROOT = Path(__file__).parent
PRACTICES = 15
STUDENTS = 30
TASKS = 30
PRACTITIONERS = ["Артём", "Рами", "Немат"]

NAVY = "17365D"
PLUS_GREEN = "C6E0B4"
ONE_GREEN = "548235"
RED = "C00000"
WHITE = "FFFFFF"
INPUT = "FFFFFF"
THIN = Side(style="thin", color="D9E1F2")


def task_label(value):
    text = "" if value is None else str(value).strip()
    try:
        number = float(text.replace(",", "."))
        if number.is_integer():
            return str(int(number))
    except ValueError:
        pass
    return text


def hide_top_rows(ws):
    ws.row_dimensions[1].hidden = True
    ws.row_dimensions[2].hidden = True


def header_style(ws, row, end_col):
    for col in range(1, end_col + 1):
        cell = ws.cell(row, col)
        cell.font = Font(bold=True, color=WHITE)
        cell.fill = PatternFill("solid", fgColor=NAVY)
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = Border(bottom=Side(style="medium", color=WHITE))
    ws.row_dimensions[row].height = 28


def body_style(ws, first, last, first_col, last_col):
    for row in ws.iter_rows(min_row=first, max_row=last, min_col=first_col, max_col=last_col):
        for cell in row:
            cell.border = Border(bottom=THIN)
            cell.alignment = Alignment(vertical="center", wrap_text=True)


def add_validation(ws, cell_range, values, message, kind="list"):
    if kind == "list":
        validation = DataValidation(type="list", formula1='"' + values + '"', allow_blank=True)
    else:
        validation = DataValidation(type=kind, operator="between", formula1="0", formula2="4", allow_blank=True)
    validation.errorStyle = "stop"
    validation.showErrorMessage = True
    validation.errorTitle = "Недопустимое значение"
    validation.error = message
    ws.add_data_validation(validation)
    validation.add(cell_range)


def color_tasks(ws, header_row, first_row, last_row):
    for col in range(6, 36):
        ws.cell(header_row, col).fill = PatternFill("solid", fgColor="38761D")
        ws.cell(header_row, col).font = Font(bold=True, color=WHITE)
    rng = f"F{first_row}:AI{last_row}"
    anchor = f"F{first_row}"
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'{anchor}="+"'], fill=PatternFill("solid", fgColor=PLUS_GREEN), font=Font(color="000000")))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'{anchor}="1"'], fill=PatternFill("solid", fgColor=ONE_GREEN), font=Font(color=WHITE)))
    ws.conditional_formatting.add(rng, FormulaRule(
        formula=[f'{anchor}="-"'], fill=PatternFill("solid", fgColor=RED), font=Font(color=WHITE)))


def make_common(wb):
    ws = wb.create_sheet("Общий")
    hide_top_rows(ws)
    headers = ["ФИО", "Практик", "TG", "Прод", "D", "CW", "Тир", "Экзамен", "Комментарий"]
    for col, value in enumerate(headers, 1):
        ws.cell(3, col, value)
    header_style(ws, 3, 9)
    for row in range(4, 94):
        name = f"Студент{1 if row <= 33 else 2 if row <= 63 else 3}.{row - (3 if row <= 33 else 33 if row <= 63 else 63)}"
        practitioner = PRACTITIONERS[0 if row <= 33 else 1 if row <= 63 else 2]
        ws.cell(row, 1, name)
        ws.cell(row, 2, practitioner)
        ws.cell(row, 4, "Да")
        ws.cell(row, 5, f'=IF(A{row}="","",SUMIF(\'Логи\'!$A$4:$A$3006,A{row},\'Логи\'!$H$4:$H$3006))')
        ws.cell(row, 6, f'=IF(A{row}="","",SUMIF(\'CW\'!$A:$A,A{row},\'CW\'!$D:$D))')
        ws.cell(row, 7, f'=IF(A{row}="","",IF(AND(E{row}>=55,F{row}>=24),"S",IF(AND(E{row}>=35,F{row}>=14),"A",IF(E{row}>=17,"B","Допса"))))')
        for col in [1, 2, 3, 4, 8, 9]:
            ws.cell(row, col).protection = Protection(locked=False)
            ws.cell(row, col).fill = PatternFill("solid", fgColor=INPUT)
    body_style(ws, 4, 93, 1, 9)
    add_validation(ws, "B4:B93", ",".join(PRACTITIONERS), "Выберите практика")
    add_validation(ws, "D4:D93", "Да,Нет", "Выберите Да или Нет")
    add_validation(ws, "H4:H93", "неуд,уд,хорошо,очень хорошо", "Выберите категорию экзамена")
    ws.freeze_panes = "D4"
    ws.auto_filter.ref = "A3:I93"
    ws.sheet_view.showGridLines = False
    for col, width in {"A": 27, "B": 12, "C": 22, "D": 10, "E": 10, "F": 10, "G": 12, "H": 22, "I": 30}.items():
        ws.column_dimensions[col].width = width
    return ws


def make_cw(wb):
    ws = wb.create_sheet("CW")
    names = []
    for i, practitioner in enumerate(PRACTITIONERS):
        for j in range(1, STUDENTS + 1):
            names.append((f"Студент{i + 1}.{j}", practitioner))
    output = 1
    for kr in (1, 2):
        title, weights, header = output, output + 1, output + 2
        first, last = output + 3, output + 3 + len(names) - 1
        ws.cell(title, 1, f"КР-{kr}")
        ws.cell(title, 1).font = Font(size=14, bold=True, color=NAVY)
        ws.cell(title, 1).fill = PatternFill("solid", fgColor="D9D9D9")
        ws.cell(weights, 4, "W")
        for col in range(5, 10):
            ws.cell(weights, col, 4)
        for col, value in enumerate(["ФИО", "Попытка", "TG", "S", "1", "2", "3", "4", "5"], 1):
            ws.cell(header, col, value)
        header_style(ws, header, 9)
        for row, (name, _practitioner) in enumerate(names, first):
            ws.cell(row, 1, name)
            ws.cell(row, 2).fill = PatternFill("solid", fgColor="FFF2CC")
            ws.cell(row, 2).protection = Protection(locked=False)
            ws.cell(row, 4, f'=IF(A{row}="","",SUMPRODUCT(E{row}:I{row},E${weights}:I${weights}))')
            for col in range(5, 10):
                ws.cell(row, col).protection = Protection(locked=False)
                ws.cell(row, col).fill = PatternFill("solid", fgColor=INPUT)
        body_style(ws, first, last, 1, 9)
        add_validation(ws, f"B{first}:B{last}", "1,2,3,4", "Введите целое число от 1 до 4")
        ws.data_validations.dataValidation[-1].type = "whole"
        ws.data_validations.dataValidation[-1].operator = "between"
        ws.data_validations.dataValidation[-1].formula1 = "1"
        ws.data_validations.dataValidation[-1].formula2 = "4"
        score = DataValidation(type="decimal", operator="between", formula1="0", formula2="4", allow_blank=True)
        score.errorStyle = "stop"
        score.showErrorMessage = True
        score.error = "Введите число от 0 до 4"
        ws.add_data_validation(score)
        score.add(f"E{first}:I{last}")
        output = last + 2
    ws.freeze_panes = "E4"
    ws.sheet_view.showGridLines = False
    for col, width in {"A": 27, "B": 12, "C": 22, "D": 12}.items():
        ws.column_dimensions[col].width = width
    for col in range(5, 10):
        ws.column_dimensions[get_column_letter(col)].width = 11
    return ws


def make_ranking(wb):
    ws = wb.create_sheet("Рейтинг")
    hide_top_rows(ws)
    for col, value in enumerate(["ФИО", "Практик", "TG", "D", "CW", "Тир"], 1):
        ws.cell(3, col, value)
    header_style(ws, 3, 6)
    body_style(ws, 4, 93, 1, 6)
    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = False
    for col, width in {"A": 27, "B": 12, "C": 22, "D": 10, "E": 10, "F": 12}.items():
        ws.column_dimensions[col].width = width
    return ws


def make_logs(wb):
    ws = wb.create_sheet("Логи")
    hide_top_rows(ws)
    headers = ["ФИО", "Практик", "Практика", "Задача", "Номер выхода", "База D", "Коэффициент", "Начислено D", "Статус", "Поправка", "Дата"]
    for col, value in enumerate(headers, 1):
        ws.cell(3, col, value)
    header_style(ws, 3, 11)
    for row in range(4, 3007):
        ws.cell(row, 5, f'=IF(A{row}="","",COUNTIFS($A$4:A{row},A{row},$F$4:F{row},"<>",$I$4:I{row},"<>ушел на базу"))')
        ws.cell(row, 6, f'=IF(E{row}="","",IF(E{row}<=3,6,IF(E{row}<=6,5,IF(E{row}<=9,4,IF(E{row}<=12,3,IF(E{row}<=15,2,0))))))')
        ws.cell(row, 8, f'=IF(OR(F{row}="",G{row}=""),"",F{row}*G{row}+IF(J{row}="",0,J{row}))')
        for col in [1, 2, 3, 4, 7, 9, 10, 11]:
            ws.cell(row, col).protection = Protection(locked=False)
    body_style(ws, 4, 3006, 1, 11)
    ws.freeze_panes = "A4"
    ws.sheet_view.showGridLines = False
    for col, width in {"A": 27, "B": 12, "C": 12, "D": 14, "E": 14, "F": 10, "G": 14, "H": 14, "I": 28, "J": 12, "K": 18}.items():
        ws.column_dimensions[col].width = width
    return ws


def make_plus(path, practitioner, index):
    wb = Workbook()
    ws = wb.active
    ws.title = "Практика 1"
    roster = [(f"Студент{index}.{i}", "") for i in range(1, STUDENTS + 1)]
    labels = [str(i) for i in range(1, TASKS + 1)]
    for practice in range(1, PRACTICES + 1):
        if practice > 1:
            ws = wb.create_sheet(f"Практика {practice}")
        header = 3
        first, last = 4, 3 + STUDENTS
        hide_top_rows(ws)
        ws.cell(1, 1, f"Практика {practice} — плюсы")
        ws.cell(1, 1).font = Font(size=14, bold=True, color=NAVY)
        ws.cell(1, 1).fill = PatternFill("solid", fgColor="D9D9D9")
        ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=35)
        for col, value in enumerate(["ФИО", "TG", "S", "Коэффициент", "D за практику"] + labels, 1):
            ws.cell(header, col, value)
        header_style(ws, header, 35)
        for col, fill in [(1, "A9D18E"), (2, "9DC3E6"), (3, "F4B183"), (4, "FFF2CC"), (5, "D9EAD3")]:
            ws.cell(header, col).fill = PatternFill("solid", fgColor=fill)
            ws.cell(header, col).font = Font(bold=True, color="000000")
        for row, (name, tg) in enumerate(roster, first):
            ws.cell(row, 1, name)
            ws.cell(row, 2, tg)
            ws.cell(row, 3, f'=IF(A{row}="","",COUNTIF(F{row}:AI{row},"+")+COUNTIF(F{row}:AI{row},"1"))')
            ws.cell(row, 4, f'=IF(A{row}="","",IF(C{row}>2*COUNTA(F${header}:AI${header})/3,1,IF(C{row}>=COUNTA(F${header}:AI${header})/2,0.8,0.5)))')
            for col in range(6, 36):
                ws.cell(row, col).number_format = "@"
                ws.cell(row, col).protection = Protection(locked=False)
                ws.cell(row, col).fill = PatternFill("solid", fgColor=INPUT)
        body_style(ws, first, last, 1, 35)
        color_tasks(ws, header, first, last)
        ws.conditional_formatting.add(f"C{first}:C{last}", CellIsRule(operator="greaterThanOrEqual", formula=["21"], fill=PatternFill("solid", fgColor="70AD47")))
        ws.conditional_formatting.add(f"C{first}:C{last}", CellIsRule(operator="between", formula=["15", "20"], fill=PatternFill("solid", fgColor="A9D18E")))
        ws.conditional_formatting.add(f"C{first}:C{last}", CellIsRule(operator="lessThan", formula=["15"], fill=PatternFill("solid", fgColor="E2F0D9")))
        ws.freeze_panes = "F4"
        ws.sheet_view.showGridLines = False
        for col, width in {"A": 27, "B": 22, "C": 9, "D": 14, "E": 14}.items():
            ws.column_dimensions[col].width = width
        for col in range(6, 36):
            ws.column_dimensions[get_column_letter(col)].width = 10
        ws.protection.sheet = False
    wb.save(path)


def main():
    central = Workbook()
    central.remove(central.active)
    make_common(central)
    make_cw(central)
    make_ranking(central)
    make_logs(central)
    central.calculation.fullCalcOnLoad = True
    central.calculation.forceFullCalc = True
    central.calculation.calcMode = "auto"
    central.save(ROOT / "ДМобщее.xlsx")
    for index, practitioner in enumerate(PRACTITIONERS, 1):
        make_plus(ROOT / f"ДМ{practitioner}.xlsx", practitioner, index)
    print("Created four DM workbooks")


if __name__ == "__main__":
    main()
