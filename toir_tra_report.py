import re
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
import sys

# Настройка UTF-8 вывода в Windows-консоли
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass
from openpyxl.utils import get_column_letter
from openpyxl.styles import Alignment

# ============= НАСТРОЙКИ =============
# Бизнес-логика: если имя шаблона содержит "CT-GST", это означает,
# что трансмиттал предназначен для Заказчика. Это важно для будущей доработки
# при использовании разных шаблонов.
TEMPLATE_PATH = Path("Template/CT-GST-TRA-PRM-Template.xltx")  # шаблон трансмиталла (.xltx)
OUTPUT_DIR    = Path("test")                                    # куда сохранить результат .xlsx
#DOCS_DIR      = Path("test/TRA_GST")                                   # папка с файлами для таблицы
DOCS_DIR      = Path(r"D:\CT DOO\CT_docs - 01_Maintenance\03_Report_base\00_Review\04_TRA_GST\2025_T35")
TZ_FILE_PATH  = Path("Template/TZ.xlsx")                       # карта индекс -> назив

# Дата в шапке
DATE_CELL_ADDR = "C3"            # текущая дата/текстовая замена в C3
DATE_FMT_TEXT  = "%d.%m.%Y"      # dd.mm.yyyy

# Якорь футера и начало данных
FOOTER_ANCHOR_NAME = "FooterAnchor"  # именованный диапазон (у тебя B20)
FIRST_DATA_ROW     = 18              # первая строка данных

# Колонки (левая-верхняя ячейка)
COL_RB = 2   # B: "Ред. Број"
COL_BD = 3   # C: "Број документа" (объединение C:H)
COL_NZ = 9   # I: "Назив документа" (объединение I:L)

# Диапазоны объединений по строке
MERGE_BD_FROM, MERGE_BD_TO = 3, 8    # C..H
MERGE_NZ_FROM, MERGE_NZ_TO = 9, 12   # I..L

# Какие файлы включать
ALLOWED_EXT = {".pdf", ".docx", ".xlsx", ".xls", ".dwg"}

# Индекс ТЗ: II.12 / I.3.33 / II.2.7 / I.2.7a/а (+ опц. четвертый уровень)
RE_INDEX = re.compile(
    r"\b([IVXLCDM]+)\.(\d+)(?:\.(\d+))?(?:\.(\d+))?([A-Za-zА-Яа-я])?\b",
    re.IGNORECASE
)

# Поиск дат для подстановки в C3, если в ячейке текст
DATE_PATTERNS = [
    re.compile(r"\b\d{2}\.\d{2}\.\d{4}\b"),  # dd.mm.yyyy
    re.compile(r"\b\d{4}-\d{2}-\d{2}\b"),    # yyyy-mm-dd
    re.compile(r"\b\d{2}\.\d{2}\.\d{2}\b"),  # dd.mm.yy
]

# ---------- Утилиты ----------
def list_docs(doc_dir: Path):
    """Рекурсивно ищет все файлы в директории и поддиректориях."""
    return [p for p in sorted(doc_dir.rglob('*'))
            if p.is_file() and p.suffix.lower() in ALLOWED_EXT]

def write_date(ws):
    today = datetime.now().strftime(DATE_FMT_TEXT)
    cell = ws[DATE_CELL_ADDR]
    val = cell.value
    if isinstance(val, str):
        new = val
        for pat in DATE_PATTERNS:
            if pat.search(new):
                new = pat.sub(today, new, count=1)
                break
        else:
            new = today
        cell.value = new
    else:
        cell.value = today

def normalize_key(key: str) -> str:
    """
    Приводит ключ к каноническому виду: верхний регистр, кириллические суффиксы.
    """
    key = key.upper()
    # Карта замен: {латиница: кириллица}
    replacements = {
        'A': 'А',
        'B': 'Б',
        'V': 'В',
        'G': 'Г',
    }
    # Проходимся по всем заменам
    for lat, cyr in replacements.items():
        key = key.replace(lat, cyr)
    return key

def get_footer_row_by_name(wb, ws_name: str, name: str) -> int | None:
    # openpyxl >= 2.5: wb.defined_names is a DefinedNameDict
    dn = wb.defined_names.get(name)
    if dn is None:
        return None

    # dn.destinations yields (sheet_name, reference) pairs
    try:
        destinations = list(dn.destinations)
    except Exception:
        destinations = []

    for sname, ref in destinations:
        # sheet names may be quoted in the defined name
        s_clean = sname.strip("'") if isinstance(sname, str) else sname
        if s_clean == ws_name:
            coord = str(ref).split("!")[-1].replace("$", "")
            m = re.search(r"\d+", coord)
            if m:
                return int(m.group(0))
    return None

def update_footer_anchor(wb, ws_name: str, name: str, new_row: int, column_letter: str = "B") -> None:
    """
    Обновляет или создаёт именованный диапазон `name` так, чтобы он указывал на ячейку
    `<column_letter><new_row>` на листе `ws_name` (по умолчанию B<row>), например "'Лист1'!$B$25".
    Если имя не существует — создаётся. Если существует — переопределяется.
    """
    from openpyxl.workbook.defined_name import DefinedName

    ref = f"'{ws_name}'!${column_letter}${new_row}"

    # Попробуем удалить существующее определение, если оно есть
    try:
        existing = wb.defined_names.get(name)
    except Exception:
        existing = None

    # openpyxl может хранить несколько определений с одинаковым именем (по книгам/областям)
    # Для надёжности удалим все с таким именем, затем добавим наше
    try:
        wb.defined_names.delete(name)
    except Exception:
        pass

    # Добавим новое определение, совместимо с разными версиями openpyxl
    dn_obj = DefinedName(name=name, attr_text=ref)
    added = False
    # Вариант 1: словарный интерфейс
    try:
        wb.defined_names[name] = dn_obj
        added = True
    except Exception:
        pass
    # Вариант 2: метод add(name=..., attr_text=...)
    if not added:
        try:
            wb.defined_names.add(name=name, attr_text=ref)
            added = True
        except Exception:
            pass
    # Вариант 3: старый интерфейс append(DefinedName)
    if not added:
        try:
            wb.defined_names.append(dn_obj)
            added = True
        except Exception:
            pass

def ensure_space(ws, first_data_row: int, footer_row: int, need_rows: int):
    available = footer_row - first_data_row
    if need_rows > available:
        ws.insert_rows(footer_row, amount=need_rows - available)

def ensure_row_merges(ws, row, footer_row):
    """
    Готовим объединения только на строке данных `row` в зонах C..H и I..L.
    Любые merge-диапазоны, у которых max_row >= footer_row, НЕ трогаем (это футер и ниже).
    """
    from openpyxl.utils import get_column_letter

    target_cols_min = min(MERGE_BD_FROM, MERGE_NZ_FROM)
    target_cols_max = max(MERGE_BD_TO, MERGE_NZ_TO)

    # 1) Аккуратно разъединяем только те диапазоны, которые:
    #    - пересекают ТЕКУЩУЮ строку данных
    #    - лежат строго ВЫШЕ футера
    to_unmerge = []
    for mr in list(ws.merged_cells.ranges):
        min_col, min_row, max_col, max_row = mr.bounds
        if max_row >= footer_row:
            continue  # футер и ниже — не трогаем
        overlaps_row  = (min_row <= row <= max_row)
        overlaps_cols = not (max_col < target_cols_min or min_col > target_cols_max)
        if overlaps_row and overlaps_cols:
            to_unmerge.append(str(mr))
    for ref in to_unmerge:
        try:
            ws.unmerge_cells(ref)
        except Exception:
            pass

    # 2) Восстанавливаем горизонтальные объединения ровно на этой строке
    rng1 = f"{get_column_letter(MERGE_BD_FROM)}{row}:{get_column_letter(MERGE_BD_TO)}{row}"
    rng2 = f"{get_column_letter(MERGE_NZ_FROM)}{row}:{get_column_letter(MERGE_NZ_TO)}{row}"
    existing = {str(rng).replace("$", "") for rng in ws.merged_cells.ranges}
    if rng1 not in existing:
        ws.merge_cells(rng1)
    if rng2 not in existing:
        ws.merge_cells(rng2)

# ---------- Чтение TZ.xlsx: индекс -> назив ----------
def build_tz_map_from_xlsx(xlsx_path: Path) -> dict[str, str]:
    """
    Ищет в каждой строке индекс (I.3.33 / II.2.7 / I.2.7a/а).
    За 'назив' берёт колонку C (если заполнена), иначе первую содержательную справа.
    Возврат: карта UPPER(index_variant) -> naziv.
    """
    from openpyxl import load_workbook
    tz_map: dict[str, str] = {}
    if not xlsx_path.exists():
        return tz_map

    wb = load_workbook(xlsx_path, data_only=True)
    for ws in wb.worksheets:
        max_col = min(ws.max_column, 20)
        for r in range(1, ws.max_row + 1):
            # найти индекс в строке
            idx_val, idx_col = None, None
            for c in range(1, max_col + 1):
                v = ws.cell(r, c).value
                if isinstance(v, str):
                    m = RE_INDEX.search(v)
                    if m:
                        roman, num1, num2, num3, suf = m.groups()
                        suf = suf or ""
                        idx_val = f"{roman.upper()}.{num1}"
                        if num2:
                            idx_val += f".{num2}"
                        if num3:
                            idx_val += f".{num3}"
                        idx_val += suf
                        idx_col = c
                        break
            if not idx_val:
                continue

            # собрать 'назив'
            naziv = None
            vC = ws.cell(r, 3).value  # колонка C приоритетно
            if isinstance(vC, str) and vC.strip():
                naziv = vC.strip()
            else:
                for c in range((idx_col or 1) + 1, max_col + 1):
                    v = ws.cell(r, c).value
                    if isinstance(v, str) and len(v.strip()) >= 3:
                        naziv = v.strip()
                        break

            if not naziv:
                continue

            # Приводим ключ к каноническому виду и добавляем в карту
            normalized_key = normalize_key(idx_val)
            if normalized_key not in tz_map:
                tz_map[normalized_key] = naziv
    return tz_map

def extract_index_from_name(filename: str) -> str | None:
    m = RE_INDEX.search(filename)
    if not m:
        return None
    roman, num1, num2, num3, suf = m.groups()
    suf = suf or ""
    idx = f"{roman.upper()}.{num1}"
    if num2:
        idx += f".{num2}"
    if num3:
        idx += f".{num3}"
    idx += suf
    return idx

# ---------- Заполнение таблицы ----------
def insert_rows_and_preserve_footer_merges(ws, insert_at_row: int, num_rows: int):
    """
    "Перемещает" футер вниз, вставляя строки и полностью сохраняя
    исходное форматирование футера (стили, высота строк, объединенные ячейки).
    """
    if num_rows <= 0:
        return

    # Диапазон колонок для копирования стиля (A..T, 1..20)
    MAX_COL_TO_COPY = 20
    
    # 1. Определяем полный диапазон футера для копирования
    footer_start_row = insert_at_row
    footer_end_row = ws.max_row
    if footer_end_row < footer_start_row:
        # Если футер пустой или не определен, просто вставляем строки
        ws.insert_rows(insert_at_row, amount=num_rows)
        return

    # 2. Копируем все данные и стили футера в память
    footer_snapshot = []
    for r_idx in range(footer_start_row, footer_end_row + 1):
        row_dim = ws.row_dimensions[r_idx]
        row_info = {
            "height": row_dim.height,
            "cells": []
        }
        for c_idx in range(1, MAX_COL_TO_COPY + 1):
            cell = ws.cell(row=r_idx, column=c_idx)
            row_info["cells"].append((cell.value, cell._style))
        footer_snapshot.append(row_info)

    # Копируем информацию об объединенных ячейках в футере
    footer_merges = []
    for mr in list(ws.merged_cells.ranges):
        if mr.min_row >= footer_start_row:
            footer_merges.append(mr)

    # 3. Временно разъединяем ячейки в старом футере
    for mr in footer_merges:
        ws.unmerge_cells(str(mr))

    # 4. Вставляем нужное количество пустых строк перед старым футером
    ws.insert_rows(insert_at_row, amount=num_rows)

    # 5. Восстанавливаем футер на новом месте
    new_footer_start_row = footer_start_row + num_rows
    for r_offset, row_info in enumerate(footer_snapshot):
        new_row_num = new_footer_start_row + r_offset
        
        # Восстанавливаем высоту строки
        if row_info["height"] is not None:
            ws.row_dimensions[new_row_num].height = row_info["height"]

        # Восстанавливаем ячейки (значение и стиль)
        for c_offset, (value, style) in enumerate(row_info["cells"]):
            col_num = 1 + c_offset
            new_cell = ws.cell(row=new_row_num, column=col_num)
            new_cell.value = value
            new_cell._style = style
            
    # 6. Восстанавливаем объединенные ячейки со сдвигом
    for mr in footer_merges:
        mr.shift(0, num_rows)
        ws.merge_cells(str(mr))


def fill_rows(ws, files, tz_map: dict, start_row: int, final_footer_row: int):
    """
    Заполняет строки данными, копируя стиль и высоту из первой строки данных (start_row)
    во все последующие. Также копирует постоянные значения из ячеек M, N, O
    и устанавливает перенос текста для ячеек C и I.
    """
    # 1. Определяем диапазон колонок для копирования стиля (B:P, 2:16)
    min_col_style = 2
    max_col_style = 16

    # 2. Запоминаем стили, высоту и постоянные значения из шаблонной строки
    template_styles = [ws.cell(row=start_row, column=j)._style for j in range(min_col_style, max_col_style + 1)]
    template_row_height = ws.row_dimensions[start_row].height
    
    const_vals = {
        13: ws.cell(row=start_row, column=13).value,
        14: ws.cell(row=start_row, column=14).value,
        15: ws.cell(row=start_row, column=15).value,
    }

    for i, p in enumerate(files, 1):
        r = start_row + i - 1
        if r >= final_footer_row:
            print(f"[WARN] Попытка записи в строку {r}, которая уже является частью футера ({final_footer_row}). Пропускается.")
            continue

        # 3. Применяем стили и высоту к новым строкам (пропуская первую)
        if r > start_row:
            if template_row_height is not None:
                ws.row_dimensions[r].height = template_row_height
            
            for j_idx, style in enumerate(template_styles):
                col = min_col_style + j_idx
                ws.cell(row=r, column=col)._style = style

        # 4. Восстанавливаем объединения для текущей строки
        ensure_row_merges(ws, r, final_footer_row)

        # 5. Заполняем ячейки данными
        # Динамические данные
        ws.cell(r, COL_RB).value = i
        
        # Број документа (C..H)
        c = ws.cell(r, COL_BD)
        c.value = p.name
        c.alignment = c.alignment + Alignment(wrap_text=True)
            
        # Назив документа (I..L)
        idx = extract_index_from_name(p.name)
        base_naziv = ""
        if idx:
            # Нормализуем ключ из имени файла и ищем его в карте
            normalized_key = normalize_key(idx)
            if normalized_key in tz_map:
                base_naziv = tz_map[normalized_key]
        
        if not base_naziv:
            base_naziv = p.stem

        # Добавляем префикс, если в имени файла есть "_CMM"
        final_naziv = base_naziv
        if "_CMM" in p.name.upper(): # Используем upper() для надежности
            final_naziv = "Листа коментара уз документ. " + base_naziv

        naziv_cell = ws.cell(r, COL_NZ)
        naziv_cell.value = final_naziv
        naziv_cell.alignment = naziv_cell.alignment + Alignment(wrap_text=True)

        # Постоянные данные
        for col_num, value in const_vals.items():
            ws.cell(row=r, column=col_num).value = value

    return final_footer_row

# ---------- Сохранение ----------
def save_with_increment(wb, out_dir: Path, prefix="CT-GST-TRA-PRM-"):
    out_dir.mkdir(parents=True, exist_ok=True)
    today = datetime.now().strftime("%y%m%d")
    n = 1
    while True:
        out = out_dir / f"{prefix}{today}_{n:02d}.xlsx"
        if not out.exists():
            wb.save(out)
            print(f"[OK] Сохранено: {out}")
            return out
        n += 1

# =============== MAIN ===============
def main():
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError(f"Шаблон не найден: {TEMPLATE_PATH}")
    if not DOCS_DIR.exists():
        raise FileNotFoundError(f"Папка с файлами не найдена: {DOCS_DIR}")
    if not TZ_FILE_PATH.exists():
        print(f"[WARN] Не найден {TZ_FILE_PATH} — 'Назив документа' будет из stem файла.")

    wb = load_workbook(TEMPLATE_PATH)  # .xltx
    ws = wb.active

    write_date(ws)

    footer_row = get_footer_row_by_name(wb, ws.title, FOOTER_ANCHOR_NAME) or 20

    files = list_docs(DOCS_DIR)

    tz_map = build_tz_map_from_xlsx(TZ_FILE_PATH)

    # --- Новая логика вставки строк ---
    num_files = len(files)
    available_data_rows = footer_row - FIRST_DATA_ROW
    
    rows_to_insert = 0
    if num_files > available_data_rows:
        rows_to_insert = num_files - available_data_rows

    if rows_to_insert > 0:
        insert_rows_and_preserve_footer_merges(ws, footer_row, rows_to_insert)

    new_footer_row = footer_row + rows_to_insert
    final_footer_row = fill_rows(ws, files, tz_map, FIRST_DATA_ROW, new_footer_row)
    update_footer_anchor(wb, ws.title, FOOTER_ANCHOR_NAME, final_footer_row)

    # --- Установка области печати ---
    # Область от B3 до колонки P и до последней заполненной строки.
    last_row = ws.max_row
    ws.print_area = f'B3:P{last_row}'
    print(f"[INFO] Область печати установлена на: {ws.print_area}")


    # 7) Сохранение результата (.xlsx)
    # ВАЖНО: т.к. книга загружена из .xltx, она помечена как шаблон.
    # Чтобы сохранить как обычный .xlsx, нужно сбросить этот флаг.
    wb.template = False
    save_with_increment(wb, OUTPUT_DIR)

if __name__ == "__main__":
    main()
