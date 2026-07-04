"""Generates the pastel budget / habit-tracker / task-planner workbook.

Run: python3 tools/generate_workbook.py
Output: Планировщик_Бюджет_Привычки.xlsx (repo root)
"""
import datetime
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.worksheet.datavalidation import DataValidation
from openpyxl.formatting.rule import CellIsRule, FormulaRule, ColorScaleRule, DataBarRule
from openpyxl.chart import PieChart, BarChart, LineChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.comments import Comment

# ---------------------------------------------------------------- palette --
PURPLE      = "D9CCE8"
PURPLE_DK   = "B9A6D6"
SAGE        = "CDE6D0"
SAGE_DK     = "9DC3A4"
SAGE_LIGHT  = "E9F4EA"
PINK        = "F7DCE2"
PINK_DK     = "E7B8C2"
PINK_LIGHT  = "FBEDF1"
YELLOW      = "F6EBC0"
YELLOW_DK   = "EAD98A"
YELLOW_LIGHT= "FBF6E3"
BLUE        = "D3E6F0"
BLUE_DK     = "A9C8DE"
BLUE_LIGHT  = "EBF3F8"
CREAM       = "FBF7EF"
PEACH       = "F3D9C4"
GRAY_TXT    = "5B5B5B"
GRAY_LIGHT  = "EDEDED"
WHITE       = "FFFFFF"

FONT_TITLE = Font(name="Roboto", size=16, bold=True, color=GRAY_TXT)
FONT_H2    = Font(name="Roboto", size=11, bold=True, color=GRAY_TXT)
FONT_LBL   = Font(name="Roboto", size=9, bold=True, color=GRAY_TXT)
FONT_BODY  = Font(name="Roboto", size=10, color=GRAY_TXT)
FONT_BIG   = Font(name="Roboto", size=18, bold=True, color=GRAY_TXT)
FONT_NOTE  = Font(name="Roboto", size=9, italic=True, color=GRAY_TXT)

THIN = Side(style="thin", color="D9D9D9")
BORDER_ALL = Border(left=THIN, right=THIN, top=THIN, bottom=THIN)

def fill(color):
    return PatternFill("solid", fgColor=color)

def style_title(ws, rng, text, color=PURPLE, font=None):
    ws.merge_cells(rng)
    cell = ws[rng.split(":")[0]]
    cell.value = text
    cell.font = font or FONT_TITLE
    cell.fill = fill(color)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws[rng]:
        for c in row:
            c.fill = fill(color)

def box_header(ws, rng, text, color):
    ws.merge_cells(rng)
    cell = ws[rng.split(":")[0]]
    cell.value = text
    cell.font = FONT_LBL
    cell.fill = fill(color)
    cell.alignment = Alignment(horizontal="center", vertical="center")
    for row in ws[rng]:
        for c in row:
            c.fill = fill(color)
            c.border = BORDER_ALL

def set_col_widths(ws, widths):
    for col, w in widths.items():
        ws.column_dimensions[col].width = w

wb = Workbook()

# =====================================================================
# SHEET: СПРАВОЧНИК (reference lists, must exist first for validations)
# =====================================================================
ref = wb.active
ref.title = "Справочник"
ref.sheet_properties.tabColor = PURPLE_DK
set_col_widths(ref, {"A": 22, "B": 22, "C": 24, "D": 16, "E": 18, "F": 14, "G": 16, "H": 24})

style_title(ref, "A1:H1", "СПРАВОЧНИК КАТЕГОРИЙ И СПИСКОВ", PURPLE)
ref.row_dimensions[1].height = 26

headers = ["Категории доходов", "Категории расходов", "Служебные (долг/вклад)",
           "Тип транзакции", "Категории задач", "Приоритет", "Статус задачи"]
for i, h in enumerate(headers):
    c = ref.cell(row=2, column=1 + i, value=h)
    c.font = FONT_LBL
    c.fill = fill(BLUE_LIGHT)
    c.border = BORDER_ALL
    c.alignment = Alignment(horizontal="center", wrap_text=True)

income_cats = ["Зарплата", "Заказы", "Подарок"]
expense_cats = ["Аренда", "Продукты", "Обучение", "Кредит", "Субподряд", "Ипотека", "Подарок", "Аутсорс"]
service_cats = ["Взял в долг", "Вернул долг", "Дал в долг", "Вернули долг", "Пополнение вклада", "Снятие со вклада"]
txn_types = ["Доход", "Расход", "Накопление", "Долг"]
task_cats = ["Работа", "Личное", "Другое"]
priorities = ["Срочно", "Высокий", "Средний", "Низкий"]
statuses = ["Не начато", "В процессе", "Готово"]

cols_data = [income_cats, expense_cats, service_cats, txn_types, task_cats, priorities, statuses]
for ci, data in enumerate(cols_data):
    for ri, val in enumerate(data):
        cell = ref.cell(row=3 + ri, column=1 + ci, value=val)
        cell.font = FONT_BODY
        cell.border = BORDER_ALL

note_row = 3 + max(len(c) for c in cols_data) + 1
ref.merge_cells(start_row=note_row, start_column=3, end_row=note_row + 3, end_column=3)
note_cell = ref.cell(row=note_row, column=3)
note_cell.value = ("⚠ Эти категории используются в формулах расчёта долгов и накоплений.\n"
                    "Не изменяйте и не удаляйте их, если хотите корректный учёт.\n\n"
                    "Если вы не планируете отслеживать долги и накопления — эти категории можно удалить.")
note_cell.font = FONT_NOTE
note_cell.fill = fill(YELLOW_LIGHT)
note_cell.alignment = Alignment(wrap_text=True, vertical="top")

# combined category list used by the "Категория" dropdown on Бюджет sheet
ref["J2"] = "Все категории (для выпадающего списка)"
ref["J2"].font = FONT_LBL
all_cats = income_cats + [c for c in expense_cats if c not in income_cats] + service_cats
for i, val in enumerate(all_cats):
    ref.cell(row=3 + i, column=10, value=val).font = FONT_BODY
ref.column_dimensions["J"].width = 26

ref.sheet_view.showGridLines = False

def named_list_range(col_letter, n):
    return f"Справочник!${col_letter}$3:${col_letter}${2+n}"

RANGE_INCOME = named_list_range("A", len(income_cats))
RANGE_EXPENSE = named_list_range("B", len(expense_cats))
RANGE_ALLCATS = named_list_range("J", len(all_cats))
RANGE_TXNTYPE = named_list_range("D", len(txn_types))
RANGE_TASKCAT = named_list_range("E", len(task_cats))
RANGE_PRIORITY = named_list_range("F", len(priorities))
RANGE_STATUS = named_list_range("G", len(statuses))

def add_checkbox_validation(ws, cell_range):
    dv = DataValidation(type="list", formula1='"TRUE,FALSE"', allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(cell_range)
    return dv

def add_list_validation(ws, cell_range, source_range):
    dv = DataValidation(type="list", formula1=f"={source_range}", allow_blank=True)
    ws.add_data_validation(dv)
    dv.add(cell_range)
    return dv

# =====================================================================
# SHEET: ТРЕКЕР ПРИВЫЧЕК
# =====================================================================
hab = wb.create_sheet("Трекер привычек")
hab.sheet_properties.tabColor = SAGE_DK
hab.sheet_view.showGridLines = False

N_DAYS = 31
FIRST_DAY_COL = 3  # column C
LAST_DAY_COL = FIRST_DAY_COL + N_DAYS - 1  # AG = 33

widths = {"A": 3, "B": 27}
for col in range(FIRST_DAY_COL, LAST_DAY_COL + 1):
    widths[get_column_letter(col)] = 3.4
set_col_widths(hab, widths)

# Row1: title bar
style_title(hab, f"A1:{get_column_letter(LAST_DAY_COL)}1", "🌿  ТРЕКЕР ПРИВЫЧЕК", PURPLE)
hab.row_dimensions[1].height = 28

# Row2: month-start control
hab["A2"] = "Дата начала месяца:"
hab["A2"].font = FONT_LBL
hab.merge_cells("A2:A2")
hab["B2"] = datetime.date(2026, 7, 1)
hab["B2"].number_format = "dd.mm.yyyy"
hab["B2"].font = Font(name="Roboto", size=10, bold=True, color=GRAY_TXT)
hab["B2"].fill = fill(YELLOW_LIGHT)
hab["B2"].alignment = Alignment(horizontal="center")
hab["B2"].border = BORDER_ALL
START_CELL = "$B$2"

# Row4-5: month/date pink block (merged 2 rows)
hab.merge_cells("B4:B5")
month_cell = hab["B4"]
month_cell.value = f'=UPPER(TEXT({START_CELL},"mmmm"))&CHAR(10)&TEXT({START_CELL},"d mmmm yyyy")'
month_cell.font = Font(name="Roboto", size=13, bold=True, color=GRAY_TXT)
month_cell.fill = fill(PINK)
month_cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
for rr in (4, 5):
    hab.cell(row=rr, column=2).border = BORDER_ALL

week_blocks = [
    ("НЕДЕЛЯ 1", FIRST_DAY_COL, FIRST_DAY_COL + 6, BLUE),
    ("НЕДЕЛЯ 2", FIRST_DAY_COL + 7, FIRST_DAY_COL + 13, SAGE),
    ("НЕДЕЛЯ 3", FIRST_DAY_COL + 14, FIRST_DAY_COL + 20, PINK),
    ("НЕДЕЛЯ 4", FIRST_DAY_COL + 21, FIRST_DAY_COL + 27, YELLOW),
    ("ДОП.", FIRST_DAY_COL + 28, LAST_DAY_COL, SAGE),
]
for label, c1, c2, color in week_blocks:
    rng = f"{get_column_letter(c1)}4:{get_column_letter(c2)}4"
    box_header(hab, rng, label, color)

# Row5: weekday abbreviations, Row6: date numbers
for col in range(FIRST_DAY_COL, LAST_DAY_COL + 1):
    idx = col - FIRST_DAY_COL
    letter = get_column_letter(col)
    wd = hab.cell(row=5, column=col)
    wd.value = f'=TEXT({START_CELL}+{idx},"ddd")'
    wd.font = FONT_LBL
    wd.fill = fill(BLUE_LIGHT)
    wd.alignment = Alignment(horizontal="center")
    wd.border = BORDER_ALL
    dn = hab.cell(row=6, column=col)
    dn.value = f'=DAY({START_CELL}+{idx})'
    dn.font = Font(name="Roboto", size=9, bold=True, color=GRAY_TXT)
    dn.fill = fill(CREAM)
    dn.alignment = Alignment(horizontal="center")
    dn.border = BORDER_ALL

# Habit rows 7-21 (15 rows)
HAB_FIRST_ROW, HAB_LAST_ROW = 7, 21
example_habits = [
    "💧 2 литра воды", "🏃 Спортзал в 19:00", "📖 Чтение 30 мин", "📓 Ведение дневника",
    "💼 Работа — 5 часов", "💪 Подработка — 2 часа", "🧘 Растяжка 15 мин", "🙏 Практика благодарности",
]
for i, r in enumerate(range(HAB_FIRST_ROW, HAB_LAST_ROW + 1)):
    hab.cell(row=r, column=1, value=i + 1).font = FONT_BODY
    name_cell = hab.cell(row=r, column=2)
    if i < len(example_habits):
        name_cell.value = example_habits[i]
    name_cell.font = FONT_BODY
    name_cell.border = BORDER_ALL
    day_range = f"{get_column_letter(FIRST_DAY_COL)}{r}:{get_column_letter(LAST_DAY_COL)}{r}"
    for col in range(FIRST_DAY_COL, LAST_DAY_COL + 1):
        c = hab.cell(row=r, column=col, value=False)
        c.alignment = Alignment(horizontal="center")
        c.border = BORDER_ALL
    add_checkbox_validation(hab, day_range)
    hab.conditional_formatting.add(
        day_range,
        FormulaRule(formula=[f"{get_column_letter(FIRST_DAY_COL)}{r}=TRUE"], fill=fill(SAGE))
    )

NAME_RANGE = f"$B${HAB_FIRST_ROW}:$B${HAB_LAST_ROW}"

# Summary rows: Выполнены / Невыполнены / Прогресс %
ROW_DONE, ROW_NOTDONE, ROW_PCT = HAB_LAST_ROW + 1, HAB_LAST_ROW + 2, HAB_LAST_ROW + 3
hab.cell(row=ROW_DONE, column=2, value="ВЫПОЛНЕНЫ").font = FONT_LBL
hab.cell(row=ROW_NOTDONE, column=2, value="НЕВЫПОЛНЕНЫ").font = FONT_LBL
hab.cell(row=ROW_PCT, column=2, value="ПРОГРЕСС %").font = FONT_LBL
for r in (ROW_DONE, ROW_NOTDONE, ROW_PCT):
    hab.cell(row=r, column=2).fill = fill(GRAY_LIGHT)
    hab.cell(row=r, column=2).border = BORDER_ALL

for col in range(FIRST_DAY_COL, LAST_DAY_COL + 1):
    letter = get_column_letter(col)
    d = hab.cell(row=ROW_DONE, column=col)
    d.value = f'=COUNTIFS({NAME_RANGE},"<>",{letter}{HAB_FIRST_ROW}:{letter}{HAB_LAST_ROW},TRUE)'
    nd = hab.cell(row=ROW_NOTDONE, column=col)
    nd.value = f'=COUNTIFS({NAME_RANGE},"<>",{letter}{HAB_FIRST_ROW}:{letter}{HAB_LAST_ROW},FALSE)'
    p = hab.cell(row=ROW_PCT, column=col)
    p.value = f'=IFERROR({letter}{ROW_DONE}/({letter}{ROW_DONE}+{letter}{ROW_NOTDONE}),0)'
    p.number_format = "0%"
    for cell in (d, nd, p):
        cell.font = Font(name="Roboto", size=9, color=GRAY_TXT)
        cell.alignment = Alignment(horizontal="center")
        cell.border = BORDER_ALL

pct_range = f"{get_column_letter(FIRST_DAY_COL)}{ROW_PCT}:{get_column_letter(LAST_DAY_COL)}{ROW_PCT}"
hab.conditional_formatting.add(
    pct_range,
    ColorScaleRule(start_type="min", start_color="FBEDF1", end_type="max", end_color=SAGE_DK)
)

hab.freeze_panes = f"C{HAB_FIRST_ROW}"

# Weekly habits table
WEEKLY_HDR_ROW = ROW_PCT + 2
box_header(hab, f"B{WEEKLY_HDR_ROW}:B{WEEKLY_HDR_ROW}", "ЕЖЕНЕДЕЛЬНЫЕ ПРИВЫЧКИ", BLUE)
for i, wk in enumerate(["НЕДЕЛЯ 1", "НЕДЕЛЯ 2", "НЕДЕЛЯ 3", "НЕДЕЛЯ 4", "НЕДЕЛЯ 5"]):
    box_header(hab, f"{get_column_letter(3+i)}{WEEKLY_HDR_ROW}:{get_column_letter(3+i)}{WEEKLY_HDR_ROW}", wk, BLUE_LIGHT)

weekly_examples = ["🏋 Спортзал 3 раза в неделю", "💰 Анализ расходов"]
WEEKLY_FIRST, WEEKLY_LAST = WEEKLY_HDR_ROW + 1, WEEKLY_HDR_ROW + 5
for i, r in enumerate(range(WEEKLY_FIRST, WEEKLY_LAST + 1)):
    nc = hab.cell(row=r, column=2)
    if i < len(weekly_examples):
        nc.value = weekly_examples[i]
    nc.font = FONT_BODY
    nc.border = BORDER_ALL
    for col in range(3, 8):
        cc = hab.cell(row=r, column=col, value=False)
        cc.alignment = Alignment(horizontal="center")
        cc.border = BORDER_ALL
    add_checkbox_validation(hab, f"C{r}:G{r}")
    hab.conditional_formatting.add(f"C{r}:G{r}", FormulaRule(formula=[f"C{r}=TRUE"], fill=fill(SAGE)))

# Monthly habits table (to the right)
box_header(hab, f"I{WEEKLY_HDR_ROW}:J{WEEKLY_HDR_ROW}", "ЕЖЕМЕСЯЧНЫЕ ПРИВЫЧКИ", PINK)
hab.column_dimensions["I"].width = 26
hab.column_dimensions["J"].width = 8
monthly_examples = ["📏 Сделать замеры до/после"]
for i, r in enumerate(range(WEEKLY_FIRST, WEEKLY_FIRST + 2)):
    nc = hab.cell(row=r, column=9)
    if i < len(monthly_examples):
        nc.value = monthly_examples[i]
    nc.font = FONT_BODY
    nc.border = BORDER_ALL
    cc = hab.cell(row=r, column=10, value=False)
    cc.alignment = Alignment(horizontal="center")
    cc.border = BORDER_ALL
    add_checkbox_validation(hab, f"J{r}:J{r}")
    hab.conditional_formatting.add(f"J{r}:J{r}", FormulaRule(formula=[f"J{r}=TRUE"], fill=fill(PINK)))

# Notes section
NOTES_HDR_ROW = WEEKLY_LAST + 2
box_header(hab, f"B{NOTES_HDR_ROW}:P{NOTES_HDR_ROW}", "ЗАМЕТКИ", BLUE)
for r in range(NOTES_HDR_ROW + 1, NOTES_HDR_ROW + 7):
    for col in range(2, 17):
        c = hab.cell(row=r, column=col)
        c.border = Border(bottom=THIN)

# =====================================================================
# SHEET: ПЛАНИРОВЩИК ЗАДАЧ
# =====================================================================
tsk = wb.create_sheet("Планировщик задач")
tsk.sheet_properties.tabColor = BLUE_DK
tsk.sheet_view.showGridLines = False

set_col_widths(tsk, {"A": 3, "B": 32, "C": 12, "D": 8, "E": 13, "F": 7, "G": 12,
                      "H": 3, "I": 15, "J": 8, "K": 3, "L": 22, "M": 10, "N": 15, "O": 8})

style_title(tsk, "A1:G1", "🗂  МОИ ЗАДАЧИ", BLUE)
tsk.row_dimensions[1].height = 28

TASK_HDR_ROW = 3
for col, label in zip("BCDEFG", ["Задача", "Срок", "Дни ⏳", "Приоритет", "✅", "Категория"]):
    c = tsk[f"{col}{TASK_HDR_ROW}"]
    c.value = label
    c.font = FONT_LBL
    c.fill = fill(BLUE_LIGHT)
    c.alignment = Alignment(horizontal="center")
    c.border = BORDER_ALL

TASK_FIRST, TASK_LAST = TASK_HDR_ROW + 1, TASK_HDR_ROW + 12
task_names = [
    "🚶 10k шагов", "📕 20 страниц книги", "✉️ Отправить письмо клиенту", "🧘 Медитация (10 мин)",
    "💰 Посчитать бюджет на месяц", "🎬 Сделать рилс", "💳 Оплатить кредит", "🗓 План на следующую неделю",
    "📞 Позвонить маме", "🎤 Подготовить речь для конфы", "📊 Бюджет на месяц (план)", "📈 Составить отчёт",
]
task_offsets = [-3, -1, 0, 0, 2, 3, -2, 1, 0, 2, 10, -4]
task_priorities = ["Высокий", "Срочно", "Средний", "Высокий", "Низкий", "Срочно",
                    "Срочно", "Средний", "Средний", "Высокий", "Низкий", "Высокий"]
task_categories = ["Личное", "Личное", "Работа", "Личное", "Другое", "Работа",
                    "Другое", "Работа", "Личное", "Личное", "Личное", "Работа"]
task_done = [False, False, True, False, False, False, True, False, True, False, False, False]

PRIORITY_COLORS = {"Срочно": "F2B8C2", "Высокий": PURPLE, "Средний": BLUE, "Низкий": SAGE}

for i, r in enumerate(range(TASK_FIRST, TASK_LAST + 1)):
    tsk.cell(row=r, column=2, value=task_names[i]).font = FONT_BODY
    dcell = tsk.cell(row=r, column=3, value=f"=TODAY()+({task_offsets[i]})")
    dcell.number_format = "dd.mm.yyyy"
    dcell.font = FONT_BODY
    dcell.alignment = Alignment(horizontal="center")
    ddays = tsk.cell(row=r, column=4, value=f"=C{r}-TODAY()")
    ddays.font = FONT_BODY
    ddays.alignment = Alignment(horizontal="center")
    pcell = tsk.cell(row=r, column=5, value=task_priorities[i])
    pcell.font = FONT_BODY
    pcell.alignment = Alignment(horizontal="center")
    fcell = tsk.cell(row=r, column=6, value=task_done[i])
    fcell.alignment = Alignment(horizontal="center")
    ccell = tsk.cell(row=r, column=7, value=task_categories[i])
    ccell.font = FONT_BODY
    ccell.alignment = Alignment(horizontal="center")
    for col in range(2, 8):
        tsk.cell(row=r, column=col).border = BORDER_ALL

TASK_RANGE_ALL = f"B{TASK_FIRST}:G{TASK_LAST}"
add_list_validation(tsk, f"E{TASK_FIRST}:E{TASK_LAST}", RANGE_PRIORITY)
add_list_validation(tsk, f"G{TASK_FIRST}:G{TASK_LAST}", RANGE_TASKCAT)
add_checkbox_validation(tsk, f"F{TASK_FIRST}:F{TASK_LAST}")

for prio, color in PRIORITY_COLORS.items():
    tsk.conditional_formatting.add(
        f"E{TASK_FIRST}:E{TASK_LAST}",
        FormulaRule(formula=[f'E{TASK_FIRST}="{prio}"'], fill=fill(color))
    )
tsk.conditional_formatting.add(
    TASK_RANGE_ALL,
    FormulaRule(formula=[f"$F{TASK_FIRST}=TRUE"], font=Font(name="Roboto", size=10, strike=True, color="ABABAB"))
)
tsk.conditional_formatting.add(
    f"D{TASK_FIRST}:D{TASK_LAST}",
    CellIsRule(operator="lessThan", formula=["0"], font=Font(name="Roboto", size=10, bold=True, color="C0392B"))
)

# ---- summary stat rows ----
STAT_ROW1 = TASK_LAST + 2
prio_labels = ["СРОЧНЫЕ", "ВЫСОКИЙ ПРИОРИТЕТ", "СРЕДНИЙ ПРИОРИТЕТ", "НИЗКИЙ ПРИОРИТЕТ"]
prio_colors_stat = ["F2B8C2", PURPLE, BLUE, SAGE]
for i, (lbl, color) in enumerate(zip(prio_labels, prio_colors_stat)):
    r = STAT_ROW1 + i
    lc = tsk.cell(row=r, column=2, value=lbl)
    lc.font = FONT_LBL
    lc.fill = fill(color)
    lc.border = BORDER_ALL
    prio_word = ["Срочно", "Высокий", "Средний", "Низкий"][i]
    vc = tsk.cell(row=r, column=3, value=f'=COUNTIF($E${TASK_FIRST}:$E${TASK_LAST},"{prio_word}")')
    vc.font = FONT_H2
    vc.fill = fill(color)
    vc.alignment = Alignment(horizontal="center")
    vc.border = BORDER_ALL

cat_labels = ["ВСЕГО ЗАДАЧ", "ЛИЧНЫЕ", "РАБОТА", "ДРУГИЕ"]
for i, lbl in enumerate(cat_labels):
    r = STAT_ROW1 + i
    tsk.merge_cells(f"E{r}:F{r}")
    lc = tsk.cell(row=r, column=5, value=lbl)
    lc.font = FONT_LBL
    lc.fill = fill(BLUE_LIGHT)
    lc.alignment = Alignment(horizontal="center")
    vc = tsk.cell(row=r, column=7)
    if i == 0:
        vc.value = f"=COUNTA($B${TASK_FIRST}:$B${TASK_LAST})"
    else:
        word = {1: "Личное", 2: "Работа", 3: "Другое"}[i]
        vc.value = f'=COUNTIF($G${TASK_FIRST}:$G${TASK_LAST},"{word}")'
    vc.font = FONT_H2
    vc.fill = fill(BLUE_LIGHT)
    vc.alignment = Alignment(horizontal="center")
    for cc in (lc, vc):
        cc.border = BORDER_ALL

# filter control
FILTER_ROW = STAT_ROW1 + 5
tsk.cell(row=FILTER_ROW, column=2, value="Выделение задач:").font = FONT_LBL
filt_cell = tsk.cell(row=FILTER_ROW, column=3, value="(Все)")
filt_cell.fill = fill(YELLOW_LIGHT)
filt_cell.border = BORDER_ALL
filt_cell.alignment = Alignment(horizontal="center")
dv_filter = DataValidation(type="list", formula1='"(Все),Работа,Личное,Другое"', allow_blank=True)
tsk.add_data_validation(dv_filter)
dv_filter.add(f"C{FILTER_ROW}")
tsk.conditional_formatting.add(
    TASK_RANGE_ALL,
    FormulaRule(formula=[f'AND($C${FILTER_ROW}<>"(Все)",$G{TASK_FIRST}=$C${FILTER_ROW})'], fill=fill(YELLOW_LIGHT))
)

# stat tiles: done/progress/deadline-today/overdue
TILE_ROW = STAT_ROW1
tsk.cell(row=TILE_ROW, column=12, value="ЗАДАЧ ВЫПОЛНЕНО").font = FONT_LBL
tile_done = tsk.cell(row=TILE_ROW, column=13,
                      value=f'=COUNTIF($F${TASK_FIRST}:$F${TASK_LAST},TRUE)&"/"&COUNTA($B${TASK_FIRST}:$B${TASK_LAST})')
tsk.cell(row=TILE_ROW + 1, column=12, value="ПРОГРЕСС").font = FONT_LBL
tile_pct = tsk.cell(row=TILE_ROW + 1, column=13,
                     value=f'=COUNTIF($F${TASK_FIRST}:$F${TASK_LAST},TRUE)/COUNTA($B${TASK_FIRST}:$B${TASK_LAST})')
tile_pct.number_format = "0%"
tsk.cell(row=TILE_ROW + 2, column=12, value="ДЕДЛАЙН СЕГОДНЯ").font = FONT_LBL
tsk.cell(row=TILE_ROW + 2, column=13, value=f'=COUNTIF($C${TASK_FIRST}:$C${TASK_LAST},TODAY())')
tsk.cell(row=TILE_ROW + 3, column=12, value="ПРОСРОЧЕННЫЕ ЗАДАЧИ").font = FONT_LBL
tile_over = tsk.cell(row=TILE_ROW + 3, column=13,
                      value=f'=COUNTIFS($C${TASK_FIRST}:$C${TASK_LAST},"<"&TODAY(),$F${TASK_FIRST}:$F${TASK_LAST},FALSE)')
for i in range(4):
    lbl_c = tsk.cell(row=TILE_ROW + i, column=12)
    val_c = tsk.cell(row=TILE_ROW + i, column=13)
    lbl_c.fill = fill(CREAM)
    val_c.fill = fill(CREAM)
    val_c.font = FONT_H2
    val_c.alignment = Alignment(horizontal="center")
    lbl_c.border = BORDER_ALL
    val_c.border = BORDER_ALL
tsk.conditional_formatting.add(
    f"M{TILE_ROW+3}",
    CellIsRule(operator="greaterThan", formula=["0"], fill=fill("F2B8C2"), font=Font(bold=True, color="7A2530"))
)

# priority helper table (for pie chart)
PH_ROW = STAT_ROW1
for i, (prio, color) in enumerate(PRIORITY_COLORS.items()):
    r = PH_ROW + i
    tsk.cell(row=r, column=9, value=prio).font = FONT_BODY
    tsk.cell(row=r, column=10, value=f'=COUNTIF($E${TASK_FIRST}:$E${TASK_LAST},"{prio}")').font = FONT_BODY

# category helper table (for bar chart)
CH_ROW = STAT_ROW1
for i, word in enumerate(["Работа", "Личное", "Другое"]):
    r = CH_ROW + i
    tsk.cell(row=r, column=14, value=word).font = FONT_BODY
    tsk.cell(row=r, column=15, value=f'=COUNTIF($G${TASK_FIRST}:$G${TASK_LAST},"{word}")').font = FONT_BODY

pie = PieChart()
pie.title = "Задачи по приоритету"
data = Reference(tsk, min_col=10, min_row=PH_ROW, max_row=PH_ROW + 3)
cats = Reference(tsk, min_col=9, min_row=PH_ROW, max_row=PH_ROW + 3)
pie.add_data(data, titles_from_data=False)
pie.set_categories(cats)
pie.height, pie.width = 7, 9
pie.dataLabels = DataLabelList()
pie.dataLabels.showPercent = True
tsk.add_chart(pie, f"I{FILTER_ROW+1}")

bar = BarChart()
bar.type = "col"
bar.title = "Задачи по категориям"
bdata = Reference(tsk, min_col=15, min_row=CH_ROW, max_row=CH_ROW + 2)
bcats = Reference(tsk, min_col=14, min_row=CH_ROW, max_row=CH_ROW + 2)
bar.add_data(bdata, titles_from_data=False)
bar.set_categories(bcats)
bar.height, bar.width = 7, 9
bar.legend = None
tsk.add_chart(bar, f"N{FILTER_ROW+1}")

# ---- FILTER tables by priority (2x2 grid) ----
FT_ROW1 = FILTER_ROW + 17
filter_blocks = [
    ("СРОЧНЫЕ", "Срочно", "F2B8C2", 2),
    ("ВЫСОКИЙ ПРИОРИТЕТ", "Высокий", PURPLE, 7),
]
filter_blocks2 = [
    ("СРЕДНИЙ ПРИОРИТЕТ", "Средний", BLUE, 2),
    ("НИЗКИЙ ПРИОРИТЕТ", "Низкий", SAGE, 7),
]

def build_filter_block(ws, title, prio_word, color, start_col, header_row):
    c0 = get_column_letter(start_col)
    c3 = get_column_letter(start_col + 3)
    ws.merge_cells(f"{c0}{header_row}:{c3}{header_row}")
    hc = ws[f"{c0}{header_row}"]
    hc.value = title
    hc.font = FONT_LBL
    hc.fill = fill(color)
    hc.alignment = Alignment(horizontal="center")
    for row in ws[f"{c0}{header_row}:{c3}{header_row}"]:
        for c in row:
            c.fill = fill(color)
            c.border = BORDER_ALL
    sub_row = header_row + 1
    for i, lbl in enumerate(["Задача", "Срок", "Дни", "Тип"]):
        cc = ws.cell(row=sub_row, column=start_col + i, value=lbl)
        cc.font = FONT_LBL
        cc.fill = fill(BLUE_LIGHT if color != BLUE else BLUE_DK)
        cc.alignment = Alignment(horizontal="center")
        cc.border = BORDER_ALL
    data_row = sub_row + 1
    rng = f"$B${TASK_FIRST}:$B${TASK_LAST}"
    crit_rng = f"$E${TASK_FIRST}:$E${TASK_LAST}"
    formulas = [
        f'=IFERROR(FILTER($B${TASK_FIRST}:$B${TASK_LAST},{crit_rng}="{prio_word}"),"")',
        f'=IFERROR(FILTER($C${TASK_FIRST}:$C${TASK_LAST},{crit_rng}="{prio_word}"),"")',
        f'=IFERROR(FILTER($D${TASK_FIRST}:$D${TASK_LAST},{crit_rng}="{prio_word}"),"")',
        f'=IFERROR(FILTER($G${TASK_FIRST}:$G${TASK_LAST},{crit_rng}="{prio_word}"),"")',
    ]
    for i, f_ in enumerate(formulas):
        cell = ws.cell(row=data_row, column=start_col + i, value=f_)
        cell.font = FONT_BODY
    for r in range(data_row, data_row + 4):
        for col in range(start_col, start_col + 4):
            ws.cell(row=r, column=col).border = BORDER_ALL
    return data_row + 4

for title, prio_word, color, scol in filter_blocks:
    next_row = build_filter_block(tsk, title, prio_word, color, scol, FT_ROW1)

FT_ROW2 = FT_ROW1 + 8
for title, prio_word, color, scol in filter_blocks2:
    build_filter_block(tsk, title, prio_word, color, scol, FT_ROW2)

tsk.freeze_panes = f"B{TASK_FIRST}"

# =====================================================================
# SHEET: БЮДЖЕТ (доходы и расходы)
# =====================================================================
bud = wb.create_sheet("Бюджет")
bud.sheet_properties.tabColor = SAGE_DK
bud.sheet_view.showGridLines = False

set_col_widths(bud, {"A": 3, "B": 20, "C": 12, "D": 16, "E": 20, "F": 12, "G": 3,
                      "H": 16, "I": 12, "J": 8, "K": 12, "L": 16, "M": 12, "N": 12,
                      "O": 12, "P": 10, "Q": 10, "R": 10, "S": 3})

style_title(bud, "A1:R1", "💰  БЮДЖЕТ: ДОХОДЫ И РАСХОДЫ", SAGE)
bud.row_dimensions[1].height = 28

box_header(bud, "B3:C3", "ДОЛГИ", BLUE)
box_header(bud, "E3:F3", "ВКЛАД", YELLOW)
box_header(bud, "H3:I3", "ДОХОД", SAGE)
box_header(bud, "K3:L3", "РАСХОД", PINK)
box_header(bud, "N3:O3", "ДОХОД − РАСХОД", PURPLE)

OPS_FIRST, OPS_LAST = 60, 259
D_RANGE = f"$B${OPS_FIRST}:$B${OPS_LAST}"   # Дата
T_RANGE = f"$C${OPS_FIRST}:$C${OPS_LAST}"   # Транзакция
K_RANGE = f"$D${OPS_FIRST}:$D${OPS_LAST}"   # Категория
S_RANGE = f"$E${OPS_FIRST}:$E${OPS_LAST}"   # Сумма

def sumifs_txn(txn, cat=None):
    if cat:
        return f'SUMIFS({S_RANGE},{T_RANGE},"{txn}",{K_RANGE},"{cat}")'
    return f'SUMIFS({S_RANGE},{T_RANGE},"{txn}")'

bud["B4"] = "Мой долг (нетто)"
bud["C4"] = f'=MAX({sumifs_txn("Долг","Взял в долг")}-{sumifs_txn("Долг","Вернул долг")},0)'
bud["B5"] = "Мне должны (нетто)"
bud["C5"] = f'=MAX({sumifs_txn("Долг","Дал в долг")}-{sumifs_txn("Долг","Вернули долг")},0)'
bud.merge_cells("B6:C6")
bud["B6"] = '=IF($C$4>$C$5,"🔴 Ты должен больше, чем тебе должны","🟢 Тебе должны больше, чем ты")'

bud["E4"] = "Пополнение вклада"
bud["F4"] = f'={sumifs_txn("Накопление","Пополнение вклада")}'
bud["E5"] = "Снятие со вклада"
bud["F5"] = f'={sumifs_txn("Накопление","Снятие со вклада")}'
bud["E6"] = "На вкладе"
bud["F6"] = "=F4-F5"
bud.merge_cells("E7:F7")
bud["E7"] = '=IF($F$6>=0,"🟡 Положительный баланс накоплений","🔴 Снятий больше, чем пополнений")'

for row in (4, 5):
    for col in "BE":
        bud[f"{col}{row}"].font = FONT_BODY
        bud[f"{col}{row}"].border = BORDER_ALL
    for col in "CF":
        bud[f"{col}{row}"].font = FONT_H2
        bud[f"{col}{row}"].number_format = '#,##0" ₽"'
        bud[f"{col}{row}"].alignment = Alignment(horizontal="center")
        bud[f"{col}{row}"].border = BORDER_ALL
bud["E6"].font = FONT_LBL
bud["F6"].font = FONT_H2
bud["F6"].number_format = '#,##0" ₽"'
bud["F6"].fill = fill(YELLOW_DK)
bud["B6"].font = FONT_NOTE
bud["E7"].font = FONT_NOTE

bud.merge_cells("H4:I6")
bud["H4"] = f'={sumifs_txn("Доход")}'
bud["H4"].number_format = '#,##0" ₽"'
bud["H4"].font = FONT_BIG
bud["H4"].alignment = Alignment(horizontal="center", vertical="center")
bud.merge_cells("H7:I7")
bud["H7"] = f'="Операций: "&COUNTIF({T_RANGE},"Доход")'
bud["H7"].font = FONT_NOTE
bud["H7"].alignment = Alignment(horizontal="center")

bud.merge_cells("K4:L6")
bud["K4"] = f'={sumifs_txn("Расход")}'
bud["K4"].number_format = '#,##0" ₽"'
bud["K4"].font = FONT_BIG
bud["K4"].alignment = Alignment(horizontal="center", vertical="center")
bud.merge_cells("K7:L7")
bud["K7"] = f'="Операций: "&COUNTIF({T_RANGE},"Расход")'
bud["K7"].font = FONT_NOTE
bud["K7"].alignment = Alignment(horizontal="center")

bud.merge_cells("N4:O6")
bud["N4"] = "=H4-K4"
bud["N4"].number_format = '#,##0" ₽"'
bud["N4"].font = FONT_BIG
bud["N4"].alignment = Alignment(horizontal="center", vertical="center")
bud.merge_cells("N7:O7")
bud["N7"] = '=IF($N$4>=0,"✅ Доходы превышают расходы","⚠ Расходы превышают доходы")'
bud["N7"].font = FONT_NOTE
bud["N7"].alignment = Alignment(horizontal="center")

for rng in ("H4:I6", "H7:I7", "K4:L6", "K7:L7", "N4:O6", "N7:O7"):
    for row in bud[rng]:
        for c in row:
            c.border = BORDER_ALL

# ratio helper (for pie chart)
bud["Q3"] = "Соотношение"
bud["Q3"].font = FONT_LBL
bud["Q4"], bud["R4"] = "Доход", "=H4"
bud["Q5"], bud["R5"] = "Расход", "=K4"

# category / dynamics tables
box_header(bud, "D26:F26", "ДОХОДЫ", SAGE)
box_header(bud, "H26:J26", "РАСХОДЫ", PINK)
box_header(bud, "L26:N26", "ДИНАМИКА ЗА ГОД", BLUE)
bud["O25"] = "Год:"
bud["O25"].font = FONT_LBL
bud["P25"] = 2026
bud["P25"].font = FONT_H2
bud["P25"].fill = fill(YELLOW_LIGHT)
bud["P25"].alignment = Alignment(horizontal="center")
bud["P25"].border = BORDER_ALL
YEAR_CELL = "$P$25"

for col, lbl in zip("DEF", ["Категория", "Сумма", "%"]):
    c = bud[f"{col}27"]
    c.value = lbl
    c.font = FONT_LBL
    c.fill = fill(SAGE_LIGHT)
    c.alignment = Alignment(horizontal="center")
    c.border = BORDER_ALL
for col, lbl in zip("HIJ", ["Категория", "Сумма", "%"]):
    c = bud[f"{col}27"]
    c.value = lbl
    c.font = FONT_LBL
    c.fill = fill(PINK_LIGHT)
    c.alignment = Alignment(horizontal="center")
    c.border = BORDER_ALL
for col, lbl in zip("LMN", ["Месяц", "Доход", "Расход"]):
    c = bud[f"{col}27"]
    c.value = lbl
    c.font = FONT_LBL
    c.fill = fill(BLUE_LIGHT)
    c.alignment = Alignment(horizontal="center")
    c.border = BORDER_ALL

for i, cat in enumerate(income_cats):
    r = 28 + i
    bud[f"D{r}"] = cat
    bud[f"E{r}"] = f'=SUMIFS({S_RANGE},{T_RANGE},"Доход",{K_RANGE},"{cat}")'
    bud[f"F{r}"] = f'=IFERROR(E{r}/$H$4,0)'
    bud[f"F{r}"].number_format = "0%"
    for col in "DEF":
        bud[f"{col}{r}"].font = FONT_BODY
        bud[f"{col}{r}"].border = BORDER_ALL

for i, cat in enumerate(expense_cats):
    r = 28 + i
    bud[f"H{r}"] = cat
    bud[f"I{r}"] = f'=SUMIFS({S_RANGE},{T_RANGE},"Расход",{K_RANGE},"{cat}")'
    bud[f"J{r}"] = f'=IFERROR(I{r}/$K$4,0)'
    bud[f"J{r}"].number_format = "0%"
    for col in "HIJ":
        bud[f"{col}{r}"].font = FONT_BODY
        bud[f"{col}{r}"].border = BORDER_ALL

for m in range(1, 13):
    r = 27 + m
    bud[f"L{r}"] = f'=TEXT(DATE({YEAR_CELL},{m},1),"mmm yy")'
    bud[f"M{r}"] = (f'=SUMPRODUCT(({T_RANGE}="Доход")*(YEAR({D_RANGE}+0)={YEAR_CELL})'
                     f'*(MONTH({D_RANGE}+0)={m})*{S_RANGE})')
    bud[f"N{r}"] = (f'=SUMPRODUCT(({T_RANGE}="Расход")*(YEAR({D_RANGE}+0)={YEAR_CELL})'
                     f'*(MONTH({D_RANGE}+0)={m})*{S_RANGE})')
    for col in "LMN":
        bud[f"{col}{r}"].font = FONT_BODY
        bud[f"{col}{r}"].border = BORDER_ALL

# ---- charts ----
pie_income = PieChart()
pie_income.title = "Категории доходов"
pie_income.add_data(Reference(bud, min_col=5, min_row=28, max_row=27 + len(income_cats)))
pie_income.set_categories(Reference(bud, min_col=4, min_row=28, max_row=27 + len(income_cats)))
pie_income.height, pie_income.width = 8, 10
bud.add_chart(pie_income, "B9")

pie_expense = PieChart()
pie_expense.title = "Категории расходов"
pie_expense.add_data(Reference(bud, min_col=9, min_row=28, max_row=27 + len(expense_cats)))
pie_expense.set_categories(Reference(bud, min_col=8, min_row=28, max_row=27 + len(expense_cats)))
pie_expense.height, pie_expense.width = 8, 10
bud.add_chart(pie_expense, "J9")

line = LineChart()
line.title = "Динамика: доход / расход"
line.add_data(Reference(bud, min_col=13, min_row=27, max_row=39), titles_from_data=True)
line.add_data(Reference(bud, min_col=14, min_row=27, max_row=39), titles_from_data=True)
line.set_categories(Reference(bud, min_col=12, min_row=28, max_row=39))
line.height, line.width = 8, 12
bud.add_chart(line, "R9")

pie_ratio = PieChart()
pie_ratio.title = "Соотношение"
pie_ratio.add_data(Reference(bud, min_col=18, min_row=4, max_row=5))
pie_ratio.set_categories(Reference(bud, min_col=17, min_row=4, max_row=5))
pie_ratio.height, pie_ratio.width = 8, 10
bud.add_chart(pie_ratio, "B41")

bar_month = BarChart()
bar_month.type = "bar"
bar_month.title = "По месяцам"
bar_month.add_data(Reference(bud, min_col=13, min_row=27, max_row=39), titles_from_data=True)
bar_month.add_data(Reference(bud, min_col=14, min_row=27, max_row=39), titles_from_data=True)
bar_month.set_categories(Reference(bud, min_col=12, min_row=28, max_row=39))
bar_month.height, bar_month.width = 10, 14
bud.add_chart(bar_month, "J41")

# ---- Все операции ----
box_header(bud, f"B58:F58", "ВСЕ ОПЕРАЦИИ", SAGE)
for col, lbl in zip("BCDEF", ["Дата", "Транзакция", "Категория", "Сумма", "Описание"]):
    c = bud[f"{col}59"]
    c.value = lbl
    c.font = FONT_LBL
    c.fill = fill(SAGE_LIGHT)
    c.alignment = Alignment(horizontal="center")
    c.border = BORDER_ALL

example_ops = [
    (datetime.date(2026, 1, 5), "Доход", "Зарплата", 75000, ""),
    (datetime.date(2026, 1, 6), "Расход", "Аренда", 28000, ""),
    (datetime.date(2026, 1, 8), "Расход", "Продукты", 7500, ""),
    (datetime.date(2026, 1, 10), "Накопление", "Пополнение вклада", 15000, ""),
    (datetime.date(2026, 1, 15), "Доход", "Заказы", 20000, "у друга"),
    (datetime.date(2026, 1, 18), "Расход", "Кредит", 9000, ""),
    (datetime.date(2026, 1, 20), "Долг", "Взял в долг", 10000, ""),
    (datetime.date(2026, 1, 25), "Расход", "Обучение", 10000, ""),
    (datetime.date(2026, 2, 2), "Доход", "Зарплата", 75000, ""),
    (datetime.date(2026, 2, 3), "Расход", "Аренда", 28000, ""),
    (datetime.date(2026, 2, 5), "Расход", "Продукты", 8000, ""),
    (datetime.date(2026, 2, 8), "Накопление", "Пополнение вклада", 20000, ""),
    (datetime.date(2026, 2, 12), "Доход", "Заказы", 30000, ""),
    (datetime.date(2026, 2, 15), "Долг", "Вернул долг", 5000, ""),
    (datetime.date(2026, 2, 18), "Расход", "Ипотека", 18000, ""),
]

for i, r in enumerate(range(OPS_FIRST, OPS_LAST + 1)):
    if i < len(example_ops):
        date_v, txn_v, cat_v, sum_v, desc_v = example_ops[i]
        bud[f"B{r}"] = date_v
        bud[f"B{r}"].number_format = "dd.mm.yyyy"
        bud[f"C{r}"] = txn_v
        bud[f"D{r}"] = cat_v
        bud[f"E{r}"] = sum_v
        bud[f"F{r}"] = desc_v
    for col in "BCDEF":
        bud[f"{col}{r}"].font = FONT_BODY
        bud[f"{col}{r}"].border = BORDER_ALL
    bud[f"E{r}"].number_format = '#,##0" ₽"'

add_list_validation(bud, f"C{OPS_FIRST}:C{OPS_LAST}", RANGE_TXNTYPE)
add_list_validation(bud, f"D{OPS_FIRST}:D{OPS_LAST}", RANGE_ALLCATS)

txn_colors = {"Доход": SAGE, "Расход": PINK, "Накопление": YELLOW, "Долг": PURPLE}
for txn, color in txn_colors.items():
    bud.conditional_formatting.add(
        f"C{OPS_FIRST}:C{OPS_LAST}",
        FormulaRule(formula=[f'C{OPS_FIRST}="{txn}"'], fill=fill(color))
    )

bud.freeze_panes = f"B{OPS_FIRST}"

# =====================================================================
# final touches
# =====================================================================
wb._sheets = [bud, hab, tsk, ref]
wb.active = 0

OUT_PATH = "Планировщик_Бюджет_Привычки.xlsx"
wb.save(OUT_PATH)
print(f"Saved {OUT_PATH}")
