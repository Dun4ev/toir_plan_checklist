import re
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.workbook.defined_name import DefinedName

# === НАСТРОЙКИ ===
TEMPLATE_PATH = Path("Template/CommentSheet_Template.xltx")  # или .xlsx
SEARCH_DIR = Path("test")                        # папка с .docx
REPORT_GLOB = "*.docx"                                       # фильтр отчётов
DATE_FMT = "yyyy-mm-dd"                                      # формат в Excel для D4

# Регулярка на извлечение полезных фрагментов из имени (пример)
RE_SECTION = re.compile(r"(GMS\d+)", re.IGNORECASE)
RE_I_CODE  = re.compile(r"(I\.\d+\.\d+)", re.IGNORECASE)

def ensure_named_range(ws, wb, cell, name):
    """Создаёт именованный диапазон, если ещё не существует."""
    existing = set(wb.defined_names.keys())
    if name not in existing:
        dn = DefinedName(name=name, attr_text=f"'{ws.title}'!{cell.coordinate}")
        wb.defined_names.append(dn)

def fill_basic_fields(wb, report_name: str):
    """Заполнить D1/D4 через именованные диапазоны (если их нет — пишем прямо)."""
    ws = wb.active  # или wb["ИмяЛиста"] если известно
    # Найти именованные диапазоны
    dn_map = dict(wb.defined_names.items())

    # ReportName -> D1
    if "ReportName" in dn_map:
        dests = dn_map["ReportName"].destinations
        for sheet, coord in dests:
            ws_target = wb[sheet] if isinstance(sheet, str) else sheet
            ws_target[coord].value = report_name
    else:
        ws["D1"].value = report_name
        ensure_named_range(ws, wb, ws["D1"], "ReportName")

    # CreatedDate -> D4
    today = datetime.now()
    if "CreatedDate" in dn_map:
        for sheet, coord in dn_map["CreatedDate"].destinations:
            ws_target = wb[sheet] if isinstance(sheet, str) else sheet
            cell = ws_target[coord]
            cell.value = today
            cell.number_format = DATE_FMT
    else:
        ws["D4"].value = today
        ws["D4"].number_format = DATE_FMT
        ensure_named_range(ws, wb, ws["D4"], "CreatedDate")

def fill_extra_fields(wb, report_name: str):
    """
    Пример расширения: извлечь GMS* и I.*.* и записать в ExtraField1 (D6).
    Настройте под свои правила. Если имени нет — просто пропустим.
    """
    ws = wb.active
    gms = RE_SECTION.search(report_name)
    icode = RE_I_CODE.search(report_name)

    extra = None
    if gms and icode:
        extra = f"{gms.group(1)} / {icode.group(1)}"
    elif gms:
        extra = gms.group(1)
    elif icode:
        extra = icode.group(1)

    if extra:
        # Если есть именованный диапазон ExtraField1 — используем его
        dn_map = dict(wb.defined_names.items())
        if "ExtraField1" in dn_map:
            for sheet, coord in dn_map["ExtraField1"].destinations:
                ws_target = wb[sheet] if isinstance(sheet, str) else sheet
                ws_target[coord].value = extra
        else:
            ws["D6"].value = extra  # или другая ячейка по вашему шаблону

def make_cmm_for_report(report_path: Path):
    stem = report_path.stem  # без .docx
    cmm_name = f"{stem}_CMM.xlsx"
    cmm_path = report_path.with_name(cmm_name)

    if cmm_path.exists():
        # не перезаписываем — считаем, что уже сделано
        print(f"[SKIP] Already exists: {cmm_path.name}")
        return

    wb = load_workbook(TEMPLATE_PATH)
    wb.template = False
    fill_basic_fields(wb, stem)
    fill_extra_fields(wb, stem)
    wb.save(cmm_path)
    print(f"[OK] Created: {cmm_path.name}")

def main():
    count = 0
    for docx in SEARCH_DIR.glob(REPORT_GLOB):
        # фильтр на ваши конкретные отчёты (при необходимости)
        if docx.name.startswith("CT-DR-"):
            make_cmm_for_report(docx)
            count += 1
    print(f"Done. Processed files: {count}")

if __name__ == "__main__":
    main()
