#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Генератор учебных данных для проекта «Справочник коммерческих расценок и экспертиза КП».

ВСЕ ДАННЫЕ УСЛОВНЫЕ: компании вымышлены, цены и нормативы не являются реальными.
Генератор детерминирован: повторный запуск даёт те же значения.

Версия 2: добавлен столбец «Состав цены» (работа / работа с материалом / поставка)
в ведомость, справочник, историю договоров и КП, плюс ловушки на состав цены.

Запуск:  python scripts/generate_data.py
"""
import csv
import datetime as dt
import random
from math import floor, log10
from pathlib import Path

from openpyxl import Workbook
from openpyxl.styles import Alignment, Font, PatternFill
from openpyxl.utils import get_column_letter

SEED = 2026
ROOT = Path(__file__).resolve().parent.parent
SRC = ROOT / "data" / "source"
TESTS = ROOT / "tests"
REF_DATE = dt.date(2026, 3, 31)
VAT_NEW, VAT_OLD = 0.22, 0.20          # 22% с 01.01.2026 (425-ФЗ), 20% до этого
FIXED_TS = dt.datetime(2026, 4, 1, 9, 0, 0)  # чтобы метаданные xlsx не менялись
INFLATION = 1.085                       # условный годовой рост цен

PS, VL = "ПС 110/10 кВ", "ВЛ 110 кВ"
MAT, SMR, PNR = "МТРиОб", "СМР", "ПНР"
C_SUP, C_WORK, C_WMAT = "поставка", "работа", "работа с материалом"

# (объект, наименование, формулировка №2, ед., объём, вид затрат, цена за ед. без НДС на 31.03.2026, состав цены)
_RAW = [
    # --- МТРиОб ---
    (PS, "Трансформатор силовой 110/10 кВ, 25 МВА", "Силовой трансформатор ТДН-25000/110", "шт", 2, MAT, 38_000_000, C_SUP),
    (PS, "Выключатель элегазовый 110 кВ", "Выключатель 110 кВ колонковый элегазовый", "шт", 4, MAT, 3_200_000, C_SUP),
    (PS, "Разъединитель 110 кВ", "Разъединитель горизонтально-поворотный 110 кВ", "шт", 8, MAT, 650_000, C_SUP),
    (PS, "Ограничитель перенапряжения (ОПН) 110 кВ", "ОПН-110 кВ", "шт", 6, MAT, 120_000, C_SUP),
    (PS, "Трансформатор тока 110 кВ", "ТТ 110 кВ элегазовый", "шт", 6, MAT, 280_000, C_SUP),
    (PS, "Трансформатор напряжения 110 кВ", "ТН 110 кВ ёмкостный", "шт", 6, MAT, 310_000, C_SUP),
    (PS, "Ячейка КРУ 10 кВ", "Шкаф КРУ 10 кВ с вакуумным выключателем", "шт", 24, MAT, 1_450_000, C_SUP),
    (PS, "Трансформатор собственных нужд 10/0,4 кВ, 250 кВА", "ТСН ТМГ-250/10", "шт", 2, MAT, 620_000, C_SUP),
    (PS, "Панель релейной защиты и автоматики", "Шкаф РЗА (комплект)", "шт", 12, MAT, 540_000, C_SUP),
    (PS, "Кабель силовой 10 кВ АПвПу 3x240", "Кабель АПвПу-10 3х240", "м", 1800, MAT, 4_200, C_SUP),
    (PS, "Кабель контрольный КВВГ", "Контрольный кабель КВВГнг", "м", 9000, MAT, 95, C_SUP),
    (VL, "Провод неизолированный АС 240/32", "Провод АС-240/32", "т", 34, MAT, 410_000, C_SUP),
    (VL, "Опора стальная многогранная ВЛ 110 кВ", "Опоры ВЛ 110 кВ металлические", "т", 118, MAT, 148_000, C_SUP),
    (VL, "Изолятор подвесной стеклянный", "Изолятор ПС-120", "шт", 720, MAT, 2_100, C_SUP),
    (PS, "Шкаф АСУ ТП подстанции", "Комплекс АСУ ТП", "шт", 2, MAT, 3_800_000, C_SUP),
    (PS, "Аккумуляторная батарея с зарядным устройством", "АКБ с ЗПУ", "компл", 1, MAT, 2_100_000, C_SUP),
    # --- СМР ---
    (PS, "Разработка грунта экскаватором", "Земляные работы: разработка грунта", "м3", 4200, SMR, 380, C_WORK),
    (PS, "Устройство песчаной подушки", "Песчаное основание", "м3", 1100, SMR, 1_650, C_WMAT),
    (PS, "Бетонирование фундаментов B25", "Устройство монолитных фундаментов B25", "м3", 960, SMR, 9_800, C_WMAT),
    (PS, "Армирование фундаментов", "Арматурные работы (фундаменты)", "т", 72, SMR, 118_000, C_WMAT),
    (PS, "Монтаж силового трансформатора", "Установка силового трансформатора", "шт", 2, SMR, 1_250_000, C_WORK),
    (PS, "Монтаж выключателя 110 кВ", "Установка выключателя 110 кВ", "шт", 4, SMR, 185_000, C_WORK),
    (PS, "Монтаж разъединителя 110 кВ", "Установка разъединителя 110 кВ", "шт", 8, SMR, 78_000, C_WORK),
    (PS, "Монтаж ОПН 110 кВ", "Установка ОПН 110 кВ", "шт", 6, SMR, 24_000, C_WORK),
    (PS, "Монтаж трансформаторов тока и напряжения 110 кВ", "Установка ТТ и ТН 110 кВ", "шт", 12, SMR, 32_000, C_WORK),
    (PS, "Монтаж ячейки КРУ 10 кВ", "Установка шкафа КРУ 10 кВ", "шт", 24, SMR, 95_000, C_WORK),
    (PS, "Монтаж панелей РЗА", "Установка шкафов РЗА", "шт", 12, SMR, 48_000, C_WORK),
    (PS, "Прокладка силового кабеля 10 кВ", "Прокладка кабеля АПвПу 10 кВ", "м", 1800, SMR, 1_150, C_WORK),
    (PS, "Прокладка контрольного кабеля", "Прокладка кабеля КВВГ", "м", 9000, SMR, 190, C_WORK),
    (PS, "Устройство заземляющего контура", "Монтаж заземляющего устройства", "м", 2400, SMR, 1_450, C_WMAT),
    (PS, "Устройство ограждения подстанции", "Ограждение территории ПС", "м", 620, SMR, 5_200, C_WMAT),
    (PS, "Устройство кабельных лотков", "Монтаж кабельных лотков", "м", 1500, SMR, 2_800, C_WMAT),
    (VL, "Монтаж опор ВЛ 110 кВ", "Сборка и установка опор ВЛ", "т", 118, SMR, 46_000, C_WORK),
    (VL, "Устройство фундаментов опор ВЛ", "Фундаменты под опоры ВЛ", "м3", 320, SMR, 14_500, C_WMAT),
    (VL, "Монтаж провода ВЛ 110 кВ (12 км x 3 фазы)", "Раскатка и натяжка провода ВЛ", "км", 36, SMR, 215_000, C_WORK),
    (VL, "Монтаж изоляторов и линейной арматуры", "Подвеска изоляторов и арматуры", "шт", 720, SMR, 1_350, C_WORK),
    (VL, "Устройство грозозащитного троса", "Монтаж грозотроса", "км", 12, SMR, 185_000, C_WMAT),
    (PS, "Благоустройство и устройство проездов", "Благоустройство территории ПС", "м2", 3200, SMR, 3_100, C_WMAT),
    # --- ПНР ---
    (PS, "ПНР силового трансформатора", "Пусконаладка силового трансформатора", "шт", 2, PNR, 420_000, C_WORK),
    (PS, "ПНР выключателей 110 кВ", "Пусконаладка выключателей 110 кВ", "шт", 4, PNR, 95_000, C_WORK),
    (PS, "ПНР ячеек КРУ 10 кВ", "Пусконаладка КРУ 10 кВ", "шт", 24, PNR, 62_000, C_WORK),
    (PS, "ПНР устройств РЗА", "Пусконаладка РЗА", "шт", 12, PNR, 88_000, C_WORK),
    (PS, "ПНР АСУ ТП подстанции", "Пусконаладка АСУ ТП", "компл", 1, PNR, 2_600_000, C_WORK),
    (PS, "Испытания кабельных линий 10 кВ", "Высоковольтные испытания кабелей 10 кВ", "шт", 6, PNR, 38_000, C_WORK),
    (PS, "Комплексное опробование и сдача подстанции", "Комплексное опробование ПС", "компл", 1, PNR, 1_900_000, C_WORK),
]
ITEMS = []
for i, r in enumerate(_RAW, 1):
    ITEMS.append(dict(code=f"P{i:02d}", obj=r[0], name=r[1], alt=r[2], unit=r[3],
                      qty=r[4], kind=r[5], price=r[6], comp=r[7]))
BY = {it["code"]: it for it in ITEMS}
assert len(ITEMS) == 45

# нормы для ресурсного метода: код -> (доля материала, доля труда, машина); остальное - машины
NORMS = {
    "P17": (0, .15, "Экскаватор 1,0 м3"), "P18": (.65, .10, "Самосвал, виброплита"),
    "P19": (.62, .15, "Автобетононасос"), "P20": (.70, .20, "Кран автомобильный 16 т"),
    "P21": (0, .35, "Кран гусеничный 100 т"), "P22": (0, .55, "Кран автомобильный 25 т"),
    "P23": (0, .55, "Кран автомобильный 25 т"), "P24": (0, .70, "Автогидроподъёмник"),
    "P25": (0, .65, "Автогидроподъёмник"), "P26": (0, .60, "Кран автомобильный 25 т"),
    "P27": (0, .80, "Автогидроподъёмник"), "P28": (0, .50, "Кабелеукладчик/кран"),
    "P29": (0, .75, "Автогидроподъёмник"), "P30": (.55, .25, "Экскаватор 1,0 м3"),
    "P31": (.60, .25, "Кран автомобильный 16 т"), "P32": (.65, .20, "Кран автомобильный 16 т"),
    "P33": (0, .40, "Кран автомобильный 50 т"), "P34": (.55, .18, "Автобетононасос"),
    "P35": (0, .45, "Трактор/лебёдка (раскатка)"), "P36": (0, .60, "Автогидроподъёмник"),
    "P37": (.60, .20, "Автогидроподъёмник"), "P38": (.60, .15, "Каток, самосвал"),
}
MACHINE_RATES = {
    "Экскаватор 1,0 м3": 3800, "Самосвал, виброплита": 3100, "Автобетононасос": 5200,
    "Кран автомобильный 16 т": 3600, "Кран гусеничный 100 т": 8400,
    "Кран автомобильный 25 т": 4300, "Автогидроподъёмник": 2900,
    "Кабелеукладчик/кран": 3900, "Кран автомобильный 50 т": 5600,
    "Трактор/лебёдка (раскатка)": 3300, "Каток, самосвал": 3400,
}
LABOR_RATE = 780       # руб./чел-ч, условно
OVERHEAD, PROFIT = 0.15, 0.07   # условные накладные и прибыль от прямых затрат

ANOMALIES = []


def anomaly(source, code, name, typ, desc, expected, origin="v1"):
    ANOMALIES.append(dict(source_file=source, item_code=code, item_name=name,
                          anomaly_type=typ, description=desc,
                          expected_conclusion=expected, origin=origin))


# ---------- утилиты ----------
def sig(x, n=4):
    """Округление до n значащих цифр (цены выглядят «живыми»)."""
    if x == 0:
        return 0
    d = n - 1 - int(floor(log10(abs(x))))
    v = round(x, d)
    return int(v) if d <= 0 else v


HEAD_FILL = PatternFill("solid", fgColor="1F3864")
F_BASE = Font(name="Arial", size=10)
F_BOLD = Font(name="Arial", size=10, bold=True)
F_HEAD = Font(name="Arial", size=10, bold=True, color="FFFFFF")
F_TITLE = Font(name="Arial", size=12, bold=True)


def new_wb():
    wb = Workbook()
    wb.properties.created = FIXED_TS
    wb.properties.modified = FIXED_TS
    wb.properties.creator = "generate_data.py"
    return wb


def write_table(ws, header, rows, title_rows=(), formats=None, widths=None,
                footer_rows=(), freeze=True):
    """Заголовочные строки, пустая строка, шапка, данные, итоги. Возвращает номер строки шапки."""
    r = 1
    for k, t in enumerate(title_rows):
        ws.cell(r, 1, t).font = F_TITLE if k == 0 else F_BASE
        r += 1
    if title_rows:
        r += 1
    hr = r
    for c, h in enumerate(header, 1):
        cell = ws.cell(r, c, h)
        cell.font, cell.fill = F_HEAD, HEAD_FILL
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    ws.row_dimensions[r].height = 32
    for row in rows:
        r += 1
        for c, v in enumerate(row, 1):
            cell = ws.cell(r, c, v)
            cell.font = F_BASE
            if formats and c in formats and isinstance(v, (int, float)):
                cell.number_format = formats[c]
    for fr in footer_rows:
        r += 1
        for c, v in enumerate(fr, 1):
            if v is None:
                continue
            cell = ws.cell(r, c, v)
            cell.font = F_BOLD
            if formats and c in formats and isinstance(v, (int, float)):
                cell.number_format = formats[c]
    if widths:
        for c, w in enumerate(widths, 1):
            ws.column_dimensions[get_column_letter(c)].width = w
    if freeze:
        ws.freeze_panes = ws.cell(hr + 1, 1)
    return hr


# ---------- 1. ведомость объёмов ----------
def gen_vor():
    wb = new_wb()
    ws = wb.active
    ws.title = "Ведомость объёмов"
    rows = [[i + 1, it["code"], it["obj"], it["kind"], it["name"], it["unit"], it["qty"], it["comp"]]
            for i, it in enumerate(ITEMS)]
    write_table(
        ws, ["№ п/п", "Код позиции", "Объект", "Вид затрат", "Наименование", "Ед. изм.", "Объём",
             "Состав цены"], rows,
        title_rows=["Ведомость объёмов работ и материалов (УЧЕБНЫЕ ДАННЫЕ)",
                    "Объект: ПС 110/10 кВ (2x25 МВА) и ВЛ 110 кВ (12 км). Заказчик вымышлен.",
                    "Состав цены: «поставка» - оборудование/материалы; «работа» - только труд и механизмы "
                    "(материал и оборудование заказчика); «работа с материалом» - материалы подрядчика."],
        formats={7: "#,##0.##"}, widths=[7, 11, 14, 11, 52, 9, 11, 22])
    wb.save(SRC / "vor_object.xlsx")


# ---------- 2. КП подрядчиков ----------
def gen_kp_A(rng):
    """Полный перечень, цены без НДС, около рынка. Состав цены - отдельной колонкой."""
    rows, total = [], 0
    for i, it in enumerate(ITEMS, 1):
        p = sig(it["price"] * rng.uniform(0.96, 1.04))
        s = round(p * it["qty"], 2)
        total += s
        rows.append([i, it["name"], it["unit"], it["qty"], p, s, it["comp"]])
    wb = new_wb()
    ws = wb.active
    ws.title = "КП"
    write_table(ws, ["№", "Наименование", "Ед. изм.", "Кол-во", "Цена за ед., руб. без НДС",
                     "Стоимость, руб. без НДС", "Состав цены"], rows,
                title_rows=["ООО «ЭнергоМонтажСервис» (вымышленная компания)",
                            "Коммерческое предложение № 41 от 12.02.2026",
                            "Цены указаны без НДС. Срок действия предложения - 90 дней."],
                formats={4: "#,##0.##", 5: "#,##0.00", 6: "#,##0.00"},
                widths=[6, 55, 9, 10, 20, 22, 22],
                footer_rows=[[None, "ИТОГО без НДС", None, None, None, round(total, 2)]])
    wb.save(SRC / "kp_A.xlsx")


def gen_kp_B(rng):
    """Цены с НДС 22%, другие формулировки, 3 пропуска, армирование +45%, ОПН -42%,
    опоры в штуках, бетонирование только как работа (в 3 раза дешевле).
    Состав цены - в названии позиции."""
    missing = {"P06", "P29", "P43"}
    rows, total, n = [], 0, 0
    for it in ITEMS:
        c = it["code"]
        if c in missing:
            continue
        n += 1
        name, unit, qty = it["alt"], it["unit"], it["qty"]
        k = rng.uniform(0.93, 1.07)
        if c == "P20":
            k *= 1.45
        if c == "P04":
            k *= 0.58
        if c == "P19":
            k *= 1 / 3.0
            name += " (работа, бетон заказчика)"
        elif it["comp"] == C_WMAT:
            name += " (с материалами)"
        price_ex = it["price"] * k
        if c == "P13":                      # опоры в штуках вместо тонн: 40 опор, ~2,95 т каждая
            unit, qty = "шт", 40
            price_ex = it["price"] * (it["qty"] / 40.0) * k
        p = sig(price_ex * (1 + VAT_NEW))
        s = round(p * qty, 2)
        total += s
        rows.append([n, name, unit, qty, p, s])
    wb = new_wb()
    ws = wb.active
    ws.title = "Лист1"
    write_table(ws, ["Поз.", "Наименование", "Ед.", "Кол-во", "Цена за ед. с НДС, руб.", "Сумма с НДС, руб."],
                rows,
                title_rows=["АО «ЭлектроСтройРесурс» (вымышленная компания)",
                            "КП № 77/26 от 20.02.2026",
                            "Все цены указаны с учётом НДС 22%."],
                formats={4: "#,##0.##", 5: "#,##0.00", 6: "#,##0.00"},
                widths=[6, 62, 8, 10, 20, 20],
                footer_rows=[[None, "ВСЕГО с НДС 22%", None, None, None, round(total, 2)]])
    wb.save(SRC / "kp_B.xlsx")


def gen_kp_V(rng):
    """Два листа, цены в тыс. руб., КП от ноября 2025 (НДС 20%, цены ниже), нет ПНР,
    силовой кабель завышен, состав цены не указан совсем, песчаная подушка подозрительно дешёвая."""
    wb = new_wb()

    def build(codes, sheet_title, first):
        ws = wb.active if first else wb.create_sheet()
        ws.title = sheet_title
        rows, total = [], 0
        for n, c in enumerate(codes, 1):
            it = BY[c]
            k = rng.uniform(0.88, 1.0) / 1.07          # более низкий уровень цен ноября 2025
            if c == "P10":
                k *= 1.38
            if c == "P18":
                k *= 0.55
            p = round(it["price"] * k * (1 + VAT_OLD) / 1000, 4)
            s = round(p * it["qty"], 3)
            total += s
            rows.append([n, it["name"], it["unit"], it["qty"], p, s])
        write_table(ws, ["№", "Наименование", "Ед.", "Объём", "Цена, тыс. руб. за ед. (с НДС 20%)",
                         "Итого, тыс. руб. (с НДС 20%)"], rows,
                    title_rows=["ООО «СеверЭнергоПроект» (вымышленная компания)",
                                "КП № 118 от 14.11.2025. Действительно до 31.12.2025.",
                                "Цены в тысячах рублей, включая НДС 20%."],
                    formats={4: "#,##0.##", 5: "#,##0.0000", 6: "#,##0.000"},
                    widths=[6, 58, 8, 10, 24, 24],
                    footer_rows=[[None, "Итого, тыс. руб.", None, None, None, round(total, 3)]])

    build([it["code"] for it in ITEMS if it["kind"] == MAT], "Оборудование", True)
    build([it["code"] for it in ITEMS if it["kind"] == SMR], "Работы", False)
    wb.save(SRC / "kp_V.xlsx")


def gen_kp_G(rng):
    """Только поставка и ПНР, цены с доставкой и НДС 22%, КРУ «под ключ», позиция «Шеф-монтаж»."""
    DELIV = 1.035
    sup = ["P01", "P02", "P03", "P04", "P05", "P06", "P08", "P09", "P15", "P16"]
    pnr = ["P39", "P40", "P42", "P43", "P44", "P45"]
    rows, total, n = [], 0, 0

    def add(name, unit, qty, price_with, what):
        nonlocal n, total
        n += 1
        p = sig(price_with)
        s = round(p * qty, 2)
        total += s
        rows.append([n, name, unit, qty, p, s, what])

    for c in sup:
        it = BY[c]
        add(it["alt"], it["unit"], it["qty"],
            it["price"] * rng.uniform(0.95, 1.06) * DELIV * (1 + VAT_NEW),
            "оборудование, доставка до объекта")
    kru = BY["P07"]["price"] + BY["P26"]["price"] + BY["P41"]["price"]
    add("КРУ 10 кВ: поставка, монтаж и ПНР (под ключ)", "ячейка", 24,
        kru * rng.uniform(0.97, 1.03) * DELIV * (1 + VAT_NEW), "под ключ, без разбивки")
    for c in pnr:
        it = BY[c]
        add(it["alt"], it["unit"], it["qty"],
            it["price"] * rng.uniform(0.95, 1.06) * (1 + VAT_NEW), "работа")
    add("Шеф-монтаж оборудования", "компл", 1, 1_150_000 * (1 + VAT_NEW), "работа")
    wb = new_wb()
    ws = wb.active
    ws.title = "КП"
    write_table(ws, ["№", "Наименование", "Ед.", "Кол-во", "Цена за ед. с доставкой и НДС 22%, руб.",
                     "Сумма с НДС, руб.", "Что входит в цену"], rows,
                title_rows=["ООО «ЭнергоКомплект Поставка» (вымышленная компания)",
                            "Коммерческое предложение исх. № 305 от 25.02.2026",
                            "Предложение на поставку оборудования и пусконаладочные работы. "
                            "Цены включают доставку до объекта и НДС 22%."],
                formats={4: "#,##0.##", 5: "#,##0.00", 6: "#,##0.00"},
                widths=[6, 52, 9, 9, 26, 20, 30],
                footer_rows=[[None, "Итого с НДС", None, None, None, round(total, 2)]])
    wb.save(SRC / "kp_G.xlsx")


# ---------- 3. история договоров ----------
CONTRACTORS = ["ООО «ВекторЭнерго»", "АО «СтройЭлектроМонтаж»", "ООО «Горизонт-Сети»",
               "ООО «ТехноПодстанция»", "АО «Магистраль Энерго»", "ООО «ЛинияПлюс»"]
REGIONS = ["Центральный ФО", "Северо-Западный ФО", "Приволжский ФО", "Уральский ФО", "Сибирский ФО"]
D0, D1 = dt.date(2023, 1, 10), dt.date(2026, 3, 25)


def gen_history(rng):
    rows = []
    for _ in range(310):
        it = rng.choice(ITEMS)
        d = D0 + dt.timedelta(days=rng.randint(0, (D1 - D0).days))
        yrs = (REF_DATE - d).days / 365.25
        price = it["price"] / (INFLATION ** yrs) * rng.uniform(0.90, 1.10)
        q = it["qty"] * rng.uniform(0.2, 1.5)
        q = max(1, round(q)) if it["unit"] in ("шт", "компл") else (round(q) if q > 20 else round(q, 1))
        rows.append(dict(date=d, contractor=rng.choice(CONTRACTORS), item=it, region=rng.choice(REGIONS),
                         name=it["name"] if rng.random() < 0.5 else it["alt"], qty=q, price=sig(price)))
    rows.sort(key=lambda r: (r["date"], r["item"]["code"]))
    outlier_idx = 150
    seq = {}
    out = []
    for i, r in enumerate(rows):
        y = r["date"].year
        seq[y] = seq.get(y, 0) + 1
        cid = f"ДГ-{y}-{seq[y]:04d}"
        price = r["price"]
        if i == outlier_idx:
            price = sig(price * 10)
            anomaly("contracts_history.xlsx", r["item"]["code"], r["name"], "Выброс в истории",
                    f"Договор {cid}: цена за единицу в 10 раз выше типичной (вероятная опечатка), "
                    f"{price} руб./{r['item']['unit']}.",
                    "Исключить из расчёта справочника как выброс или проверить источник")
        vat = VAT_NEW if r["date"] >= dt.date(2026, 1, 1) else VAT_OLD
        out.append([cid, r["date"], r["contractor"], r["item"]["obj"], r["item"]["code"], r["name"],
                    r["item"]["unit"], r["qty"], price, vat, round(price * (1 + vat), 2),
                    r["item"]["comp"], r["region"]])
    wb = new_wb()
    ws = wb.active
    ws.title = "Договоры"
    hr = write_table(ws, ["ID договора", "Дата договора", "Подрядчик", "Объект", "Код позиции",
                          "Наименование в договоре", "Ед. изм.", "Кол-во", "Цена за ед., руб. без НДС",
                          "Ставка НДС", "Цена за ед., руб. с НДС", "Состав цены", "Регион"], out,
                     formats={8: "#,##0.##", 9: "#,##0.00", 10: "0%", 11: "#,##0.00"},
                     widths=[15, 13, 26, 14, 10, 50, 9, 10, 18, 9, 18, 22, 20])
    for r in range(hr + 1, hr + 1 + len(out)):
        ws.cell(r, 2).number_format = "DD.MM.YYYY"
    wb.save(SRC / "contracts_history.xlsx")


# ---------- 4. справочник, аналоги, нормы, предел ----------
def gen_reference(rng):
    rows = []
    for it in ITEMS:
        p = sig(it["price"] * rng.uniform(0.95, 1.0))     # справочник немного «устарел»
        rows.append([it["code"], it["name"], it["unit"], it["kind"], it["comp"], p,
                     sig(p * 0.85), sig(p * 1.2), rng.randint(3, 14), REF_DATE])
    wb = new_wb()
    ws = wb.active
    ws.title = "Справочник"
    hr = write_table(ws, ["Код", "Наименование", "Ед. изм.", "Вид затрат", "Состав цены",
                          "Цена без НДС на 31.03.2026, руб.", "Мин., руб.", "Макс., руб.",
                          "Договоров в базе", "Дата актуализации"], rows,
                     title_rows=["Справочник коммерческих расценок (УЧЕБНЫЕ ДАННЫЕ)",
                                 "Уровень цен на 31.03.2026, без НДС."],
                     formats={6: "#,##0.00", 7: "#,##0.00", 8: "#,##0.00"},
                     widths=[8, 52, 9, 11, 22, 20, 14, 14, 12, 14])
    for r in range(hr + 1, hr + 1 + len(rows)):
        ws.cell(r, 10).number_format = "DD.MM.YYYY"
    wb.save(SRC / "reference_prices.xlsx")


def split_totals():
    t = {PS: 0, VL: 0}
    for it in ITEMS:
        t[it["obj"]] += it["price"] * it["qty"]
    return t


def gen_analogs(rng):
    tot = split_totals()
    ps_per_mva, vl_per_km = tot[PS] / 50, tot[VL] / 12
    spec = [("ПС 110/10 кВ «Северная»", "ПС", 110, "МВА", 32, 2022),
            ("ПС 110/10 кВ «Заводская»", "ПС", 110, "МВА", 50, 2023),
            ("ПС 110/10 кВ «Новая»", "ПС", 110, "МВА", 80, 2021),
            ("ПС 110/10 кВ «Речная»", "ПС", 110, "МВА", 25, 2024),
            ("ВЛ 110 кВ Северная - Заводская", "ВЛ", 110, "км", 8, 2022),
            ("ВЛ 110 кВ Заводская - Новая", "ВЛ", 110, "км", 15, 2023),
            ("ВЛ 110 кВ Речная - отпайка", "ВЛ", 110, "км", 20, 2025)]
    rows = []
    for i, (name, typ, kv, un, size, year) in enumerate(spec, 1):
        base = (ps_per_mva if typ == "ПС" else vl_per_km) * size * rng.uniform(0.85, 1.2)
        cost_y = base / (INFLATION ** (2026 - year))
        m = rng.uniform(0.50, 0.62) if typ == "ПС" else rng.uniform(0.38, 0.48)
        p = rng.uniform(0.04, 0.07) if typ == "ПС" else 0.0
        rows.append([f"A{i:02d}", name, typ, kv, size, un, year, rng.choice(REGIONS),
                     round(cost_y / 1e6, 2), round(INFLATION ** (2026 - year), 3),
                     round(m, 2), round(1 - m - p, 2), round(p, 2)])
    wb = new_wb()
    ws = wb.active
    ws.title = "Аналоги"
    write_table(ws, ["ID", "Объект-аналог", "Тип", "Напряжение, кВ", "Мощность / длина", "Ед.",
                     "Год ввода", "Регион", "Стоимость в ценах года ввода, млн руб. без НДС",
                     "Индекс к ценам 2026", "Доля МТРиОб", "Доля СМР", "Доля ПНР"], rows,
                title_rows=["Объекты-аналоги (УЧЕБНЫЕ ДАННЫЕ, вымышленные)"],
                formats={9: "#,##0.00", 10: "0.000", 11: "0%", 12: "0%", 13: "0%"},
                widths=[7, 36, 7, 12, 14, 7, 10, 20, 24, 14, 12, 10, 10])
    wb.save(SRC / "analog_objects.xlsx")


def gen_norms(rng):
    rows = []
    for code, (mat, lab, mach) in NORMS.items():
        it = BY[code]
        direct = it["price"] / (1 + OVERHEAD + PROFIT)
        labor_h = direct * lab / LABOR_RATE * rng.uniform(0.95, 1.05)
        mach_h = direct * (1 - mat - lab) / MACHINE_RATES[mach] * rng.uniform(0.95, 1.05)
        mat_cost = direct * mat * rng.uniform(0.95, 1.05)
        loss = {"P19": 1.5, "P20": 3.0}.get(code, 2.0) if mat else 0.0
        rows.append([code, it["name"], it["unit"], it["comp"], round(labor_h, 3), mach,
                     round(mach_h, 4), round(mat_cost, 2), loss])
    wb = new_wb()
    ws = wb.active
    ws.title = "Нормы"
    write_table(ws, ["Код", "Наименование", "Ед. изм.", "Состав цены", "Труд, чел-ч на ед.", "Машина",
                     "Машино-ч на ед.", "Материал, руб. на ед. (без потерь)", "Потери материала, %"],
                rows, title_rows=["Условные нормы труда, машин и материалов (УЧЕБНЫЕ ДАННЫЕ)"],
                formats={5: "0.000", 7: "0.0000", 8: "#,##0.00", 9: "0.0"},
                widths=[8, 52, 9, 22, 14, 28, 14, 22, 14])
    ws2 = wb.create_sheet("Ставки")
    rates = [["Труд", "Рабочий-монтажник (средний разряд)", "чел-ч", LABOR_RATE]]
    rates += [["Машина", m, "маш-ч", r] for m, r in MACHINE_RATES.items()]
    write_table(ws2, ["Тип", "Наименование", "Ед.", "Ставка, руб. без НДС"], rates,
                title_rows=["Условные ставки (УЧЕБНЫЕ ДАННЫЕ)"],
                formats={4: "#,##0.00"}, widths=[10, 40, 9, 20])
    ws3 = wb.create_sheet("Накладные")
    write_table(ws3, ["Показатель", "Значение", "База"],
                [["Накладные расходы", OVERHEAD, "от прямых затрат"],
                 ["Сметная прибыль", PROFIT, "от прямых затрат"]],
                title_rows=["Условные коэффициенты (УЧЕБНЫЕ ДАННЫЕ)"],
                formats={2: "0%"}, widths=[24, 12, 22])
    wb.save(SRC / "resource_norms.xlsx")


def gen_limit():
    tot = split_totals()
    kru = BY["P07"]["price"] + BY["P26"]["price"] + BY["P41"]["price"]
    trf = BY["P01"]["price"] + BY["P21"]["price"] + BY["P39"]["price"]
    rows = [["LIM-01", "ПС 110/10 кВ, два трансформатора, на 1 МВА установленной мощности", "руб./МВА",
             round(tot[PS] / 50 * 1.08, -3)],
            ["LIM-02", "ВЛ 110 кВ на стальных опорах, на 1 км", "руб./км", round(tot[VL] / 12 * 1.08, -3)],
            ["LIM-03", "Ячейка КРУ 10 кВ, поставка + монтаж + ПНР", "руб./ячейка", round(kru * 1.10, -3)],
            ["LIM-04", "Силовой трансформатор 110/10 кВ 25 МВА, поставка + монтаж + ПНР", "руб./шт",
             round(trf * 1.08, -3)]]
    rows = [r + ["31.03.2026", "условный норматив, значения вымышлены"] for r in rows]
    wb = new_wb()
    ws = wb.active
    ws.title = "Предел"
    write_table(ws, ["Код", "Наименование показателя", "Ед.", "Предельное значение без НДС, руб.",
                     "Уровень цен", "Примечание"], rows,
                title_rows=["УСЛОВНЫЙ предельный норматив стоимости (по аналогии с УНЦ)",
                            "ВНИМАНИЕ: значения вымышлены, к реальным нормативам отношения не имеют."],
                formats={4: "#,##0"}, widths=[9, 66, 13, 24, 13, 38])
    wb.save(SRC / "limit_conditional.xlsx")


# ---------- журнал аномалий ----------
def register_kp_anomalies():
    B, V, G = "kp_B.xlsx", "kp_V.xlsx", "kp_G.xlsx"
    anomaly(B, "*", "все позиции", "НДС в цене", "Цены указаны с НДС 22%.",
            "Привести к ценам без НДС (делить на 1,22)")
    anomaly(B, "*", "все позиции", "Другие формулировки", "Наименования отличаются от ведомости.",
            "Сопоставить по справочнику соответствий, а не по точному совпадению")
    for c in ("P06", "P29", "P43"):
        anomaly(B, c, BY[c]["name"], "Пропущенная позиция", "Позиции нет в КП.",
                "Отметить как отсутствующую, цену полного КП считать неполной")
    anomaly(B, "P20", BY["P20"]["name"], "Завышение", "Цена выше рынка примерно на 45%.",
            "Отклонение выше порога, выделить для проверки")
    anomaly(B, "P04", BY["P04"]["name"], "Занижение", "Цена ниже рынка примерно на 42%.",
            "Подозрительно низкая цена, проверить тип и производителя ОПН")
    anomaly(B, "P13", BY["P13"]["name"], "Единица измерения", "Опоры в штуках (40 шт) вместо тонн (118 т).",
            "Пересчитать через массу опоры, сравнивать только в одной единице")
    anomaly(V, "*", "все позиции", "Единицы цены", "Цены в тысячах рублей.", "Умножить на 1000")
    anomaly(V, "*", "все позиции", "Устаревшее КП и НДС",
            "КП от ноября 2025 года, НДС 20%, уровень цен ниже.",
            "Убрать НДС 20%, проиндексировать на дату сравнения")
    anomaly(V, "P39-P45", "ПНР (7 позиций)", "Пропущенный раздел", "В КП нет пусконаладочных работ.",
            "Не сравнивать итог с полными КП, ПНР добавить отдельно")
    anomaly(V, "P10", BY["P10"]["name"], "Завышение", "Цена силового кабеля выше рынка примерно на 38%.",
            "Отклонение выше порога, проверить после индексации")
    anomaly(G, "*", "все позиции", "Доставка в цене", "Цены включают доставку (+3,5%).",
            "Привести к сопоставимому составу, сравнивать без доставки")
    anomaly(G, "*", "все позиции", "НДС в цене", "Цены с НДС 22%.", "Привести к ценам без НДС")
    anomaly(G, "P07+P26+P41", "КРУ 10 кВ: поставка, монтаж и ПНР (под ключ)", "Нет прямого аналога",
            "Объединённая позиция без разбивки.",
            "Сравнивать с суммой поставки, монтажа и ПНР из справочника")
    anomaly(G, "-", "Шеф-монтаж оборудования", "Лишняя позиция", "Позиции нет в ведомости.",
            "Вынести отдельно, в сравнение не включать")
    # --- добавлено в версии 2 (состав цены) ---
    anomaly(B, "P19", BY["P19"]["name"], "Состав цены: только работа",
            "Бетонирование указано как «работа, бетон заказчика», цена примерно в 3 раза ниже рынка.",
            "Не считать подозрительно низкой: уточнить состав цены или добавить стоимость бетона",
            origin="v2_new")
    anomaly(V, "P17-P38", "все работы (СМР)", "Состав цены не указан",
            "В КП нигде не сказано, включены ли материалы.",
            "Запросить уточнение состава цены, сравнение предварительное", origin="v2_new")
    anomaly(V, "P18", BY["P18"]["name"], "Состав цены: вероятно без материала",
            "Цена примерно на 45% ниже рынка, состав не указан.",
            "Уточнить, включён ли песок: вывод «уточнить состав», а не «занижение»", origin="v2_new")


def write_anomaly_log():
    """Порядок: аномалии КП (v1), затем история, затем новые (v2)."""
    def rank(t):
        idx, a = t
        if a["origin"] == "v2_new":
            return (2, idx)
        if a["source_file"].startswith("contracts"):
            return (1, idx)
        return (0, idx)

    items = sorted(enumerate(ANOMALIES), key=rank)
    path = TESTS / "injected_anomalies.csv"
    with open(path, "w", newline="", encoding="utf-8-sig") as f:
        w = csv.writer(f, delimiter=";")
        w.writerow(["id", "source_file", "item_code", "item_name", "anomaly_type", "description",
                    "expected_conclusion", "origin"])
        for n, (_, a) in enumerate(items, 1):
            w.writerow([f"AN{n:02d}", a["source_file"], a["item_code"], a["item_name"], a["anomaly_type"],
                        a["description"], a["expected_conclusion"], a["origin"]])
    return len(items)


def main():
    SRC.mkdir(parents=True, exist_ok=True)
    TESTS.mkdir(parents=True, exist_ok=True)
    rng = random.Random(SEED)
    gen_vor()
    gen_kp_A(rng)
    gen_kp_B(rng)
    gen_kp_V(rng)
    gen_kp_G(rng)
    register_kp_anomalies()
    gen_history(rng)
    gen_reference(rng)
    gen_analogs(rng)
    gen_norms(rng)
    gen_limit()
    n = write_anomaly_log()
    print(f"Готово: 10 файлов в {SRC}, аномалий в журнале: {n}")


if __name__ == "__main__":
    main()
