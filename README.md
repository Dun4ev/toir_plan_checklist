# toir_plan_checklist
# Автогенерация листов комментариев к отчётам (`*_CMM.xlsx`)

## TL;DR
Автоматически создавать рядом с каждым отчётом `.docx` одноимённый файл `*_CMM.xlsx` из шаблона, заполняя `D1` (имя отчёта) и `D4` (дата создания). Шаблон — с именованными ячейками и таблицей комментариев. Рекомендовано: **Python (openpyxl)**; альтернативы: **VBA** и **Python + watchdog** для автосоздания при появлении файлов.

---

## Цель и контекст
- **Контекст**: отчёты имеют строгие имена, например:  
  `CT-DR-B-LP-GMS4-I.3.31-00-1M-20250806-00.docx`
- **Цель**: для каждого отчёта сформировать «лист комментариев» из шаблона, без ручного копирования, с единым форматом и валидациями.

---

## Входы и выходы
**Вход**
- Папка с отчётами `*.docx` (имена начинаются с `CT-DR-`).
- Шаблон Excel: `CommentSheet_Template.xltx` (или `.xlsx`).

**Выход**
- Для каждого отчёта создаётся соседний файл `*_CMM.xlsx` с заполненными полями.
  - Примеры:
    - `CT-DR-B-LP-GMS4-I.3.31-00-1M-20250806-00_CMM.xlsx`
    - `CT-DR-B-LP-GMS1-I.3.34-00-1M-20250804-00_CMM.xlsx`
    - `CT-DR-B-LP-GMS2-I.3.34-00-1M-20250806-00_CMM.xlsx`

---

## Правила именования и парсинг имени отчёта
- **Входной файл**: `*.docx`, имя начинается с `CT-DR-`.
- **Выходной файл**: `<ИмяОтчётаБезРасширения>_CMM.xlsx` рядом с исходным `.docx`.
- **Извлечение фрагментов из имени** (для заполнения доп. поля, опционально):
  - Участок/объект: `GMS\d+` → например, `GMS4`.
  - Код пункта: `I\.\d+\.\d+` → например, `I.3.31`.
  - Дата в имени (если нужна): `\b20\d{6}\b` → `20250806` → `2025-08-06`.

---

## Требования к шаблону Excel
Рекомендуется хранить шаблон как `.xltx` (Excel Template). При необходимости допускается `.xlsx`.

| Элемент                  | Требование                                                                     | Пример/Комментарий                      |
|-------------------------|----------------------------------------------------------------------------------|-----------------------------------------|
| Формат                  | `.xltx` (предпочтительно) или `.xlsx`                                           | `CommentSheet_Template.xltx`            |
| Именованные ячейки      | `ReportName` → D1; `CreatedDate` → D4; (опц.) `ExtraField1` → D6                | В D4 формат даты: `yyyy-mm-dd`          |
| Таблица комментариев    | Имя таблицы: `Comments`, фиксированные столбцы, автофильтры, Freeze Panes       | Удобно для печати и фильтрации          |
| Валидации               | Например, для `Статус`: `Open, In Progress, Done`                               | Data Validation                         |
| Защита                  | Лист под защитой; редактируемы — именованные поля и строки таблицы              | Пароль хранить отдельно (опц.)          |
| Печать                  | Поля страницы, повтор заголовков, предварительный просмотр                       | Для корректной PDF-печати               |

**Рекомендуемые столбцы таблицы `Comments`:**
`# | Дата | Раздел/Пункт | Комментарий | Требуемое действие | Ответ исполнителя | Статус | Ссылка/Стр.`

---

## Алгоритм
1. Найти в указанной папке все `*.docx`, имена которых начинаются с `CT-DR-`.
2. Для каждого отчёта сформировать имя `<stem>_CMM.xlsx`.
3. Если итоговый файл уже существует — пропустить (или перезаписать при режиме `--force`).
4. Открыть шаблон, создать книгу на его основе.
5. Заполнить:
   - `ReportName` (D1) = имя отчёта **без** `.docx`.
   - `CreatedDate` (D4) = текущая дата (формат `yyyy-mm-dd`).
   - (Опц.) `ExtraField1` (D6) = извлечённые фрагменты (`GMSx / I.a.b`). 
6. Сохранить книгу рядом с отчётом и записать лог.
7. По завершении вывести сводку: создано N, пропущено M, ошибок K.

**Псевдокод**
```text
for each docx in folder:
  if not name.startswith("CT-DR-"): continue
  out = f"{stem}_CMM.xlsx"
  if exists(out): log "SKIP"; continue

  wb = open(template)
  set_named("ReportName", stem)
  set_named("CreatedDate", today, "yyyy-mm-dd")
  extra = extract("GMS\d+") + " / " + extract("I\.\d+\.\d+")
  set_named("ExtraField1", extra)  # optional
  save wb as out
  log "OK", out
```

---

## Реализация: Python (рекомендовано)

### Установка
```bash
pip install openpyxl
# (опционально для автосоздания при появлении файлов)
pip install watchdog
```

### Скрипт `tools/make_cmm.py`
```python
import re
from pathlib import Path
from datetime import datetime
from openpyxl import load_workbook
from openpyxl.workbook.defined_name import DefinedName

TEMPLATE_PATH = Path("C:/Data/Templates/CommentSheet_Template.xltx")  # или .xlsx
SEARCH_DIR = Path("C:/Data/Reports")                                  # папка с .docx
DATE_FMT = "yyyy-mm-dd"

RE_SECTION = re.compile(r"(GMS\d+)", re.IGNORECASE)
RE_I_CODE  = re.compile(r"(I\.\d+\.\d+)", re.IGNORECASE)

def set_named_or_cell(wb, name, fallback_addr, value, number_format=None):
    dn_map = {dn.name: dn for dn in wb.defined_names.definedName}
    if name in dn_map:
        for sheet_name, coord in dn_map[name].destinations:
            cell = wb[sheet_name][coord]
            cell.value = value
            if number_format: cell.number_format = number_format
    else:
        ws = wb.active
        cell = ws[fallback_addr]
        cell.value = value
        if number_format: cell.number_format = number_format

def make_cmm(docx_path: Path):
    stem = docx_path.stem
    out_path = docx_path.with_name(f"{stem}_CMM.xlsx")
    if out_path.exists():
        print(f"[SKIP] {out_path.name} уже существует")
        return

    wb = load_workbook(TEMPLATE_PATH)  # .xltx или .xlsx
    set_named_or_cell(wb, "ReportName", "D1", stem)
    set_named_or_cell(wb, "CreatedDate", "D4", datetime.now(), DATE_FMT)

    # Опционально: заполнение ExtraField1 на основе имени
    gms = RE_SECTION.search(stem)
    icode = RE_I_CODE.search(stem)
    extra_parts = []
    if gms:   extra_parts.append(gms.group(1))
    if icode: extra_parts.append(icode.group(1))
    if extra_parts:
        set_named_or_cell(wb, "ExtraField1", "D6", " / ".join(extra_parts))

    wb.save(out_path)
    print(f"[OK] Создано: {out_path.name}")

def main():
    files = [p for p in SEARCH_DIR.glob("*.docx") if p.name.startswith("CT-DR-")]
    for p in files:
        make_cmm(p)
    print(f"Готово. Обработано: {len(files)}")

if __name__ == "__main__":
    main()
```

### (Опционально) Автогенерация при появлении новых файлов (`watchdog`)
```python
# tools/watch_folder.py
from watchdog.observers import Observer
from watchdog.events import FileSystemEventHandler
from pathlib import Path
import time
from make_cmm import make_cmm

WATCH_DIR = Path("C:/Data/Reports")

class Handler(FileSystemEventHandler):
    def on_created(self, event):
        p = Path(event.src_path)
        if p.is_file() and p.suffix.lower() == ".docx" and p.name.startswith("CT-DR-"):
            try:
                make_cmm(p)
            except Exception as e:
                print(f"[ERROR] {p.name}: {e}")

if __name__ == "__main__":
    observer = Observer()
    observer.schedule(Handler(), str(WATCH_DIR), recursive=False)
    observer.start()
    print(f"[WATCHING] {WATCH_DIR}")
    try:
        while True:
            time.sleep(1)
    except KeyboardInterrupt:
        observer.stop()
    observer.join()
```

---

## Реализация: VBA (альтернатива без Python)
```vba
Option Explicit

Sub MakeCMMFromFolder()
    Dim tpl As String: tpl = "C:\Data\Templates\CommentSheet_Template.xltx" ' или .xlsx
    Dim dlg As FileDialog: Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    If dlg.Show <> -1 Then Exit Sub
    Dim folderPath As String: folderPath = dlg.SelectedItems(1)

    Dim fso As Object: Set fso = CreateObject("Scripting.FileSystemObject")
    Dim file As Object, base As String, outPath As String

    Application.ScreenUpdating = False
    Dim created As Long: created = 0

    For Each file In fso.GetFolder(folderPath).Files
        If LCase(fso.GetExtensionName(file.Path)) = "docx" Then
            If Left$(fso.GetFileName(file.Path), 6) = "CT-DR-" Then
                base = fso.GetBaseName(file.Path)
                outPath = fso.BuildPath(folderPath, base & "_CMM.xlsx")
                If Not fso.FileExists(outPath) Then
                    Dim wb As Workbook: Set wb = Workbooks.Open(tpl)
                    On Error Resume Next
                    wb.Names("ReportName").RefersToRange.Value = base
                    wb.Names("CreatedDate").RefersToRange.Value = Date
                    On Error GoTo 0
                    With wb.Worksheets(1)
                        If .Range("D1").Value = "" Then .Range("D1").Value = base
                        .Range("D4").Value = Date: .Range("D4").NumberFormat = "yyyy-mm-dd"
                    End With
                    wb.SaveAs Filename:=outPath, FileFormat:=xlOpenXMLWorkbook
                    wb.Close SaveChanges:=False
                    created = created + 1
                End If
            End If
        End If
    Next file

    Application.ScreenUpdating = True
    MsgBox "Создано: " & created, vbInformation
End Sub
```

---

## Нефункциональные требования
- **Повторяемость**: одно и то же имя `.docx` всегда даёт одно и то же имя `_CMM.xlsx`.
- **Безопасность**: скрипты не изменяют исходный `.docx`; шаблон открывается для чтения, итог — в новую книгу.
- **Логи**: вывод в консоль/файл `generation.log` (время, путь, статус). 
- **Кроссплатформенность**: Python — Windows/macOS; VBA — зависит от Excel.

---

## Обработка краевых случаев
- `_CMM.xlsx` уже существует → пропуск (или режим `--force`).
- Имя не соответствует паттерну → файл создать, но `ExtraField1` пустое.
- Нет именованных диапазонов → использовать адресные ячейки D1/D4 как fallback.
- Нет прав на запись → лог ошибки, продолжить обработку.
- Дата в имени отсутствует → `CreatedDate` = сегодня.

---

## Критерии приёмки (чек-лист)
- [ ] Для каждого `CT-DR-*.docx` создан `*_CMM.xlsx` в той же папке.
- [ ] В D1 — точное имя отчёта без расширения.
- [ ] В D4 — текущая дата, формат `yyyy-mm-dd`.
- [ ] Таблица `Comments` присутствует; фильтры и Freeze Panes работают.
- [ ] (Опц.) В D6 — `GMSx / I.a.b` при наличии.
- [ ] Повторный запуск не перезаписывает готовые файлы (если не задан `--force`).

---

## Конфигурация (пример YAML для README/SPEC)
```yaml
template_path: C:/Data/Templates/CommentSheet_Template.xltx
search_dir:    C:/Data/Reports
date_format:   yyyy-mm-dd
parse:
  section_regex: "(GMS\d+)"
  icode_regex:   "(I\.\d+\.\d+)"
behavior:
  skip_if_exists: true
  force: false
log:
  enable: true
  file: C:/Data/Reports/generation.log
```

---

## Примеры соответствий
| Входной отчёт (.docx)                                   | Выходной CMM (.xlsx)                                   | D1 (ReportName)                                        | D4 (CreatedDate) | D6 (ExtraField1, опц.) |
|---------------------------------------------------------|---------------------------------------------------------|--------------------------------------------------------|------------------|------------------------|
| `CT-DR-B-LP-GMS4-I.3.31-00-1M-20250806-00.docx`         | `CT-DR-B-LP-GMS4-I.3.31-00-1M-20250806-00_CMM.xlsx`    | `CT-DR-B-LP-GMS4-I.3.31-00-1M-20250806-00`            | `yyyy-mm-dd`     | `GMS4 / I.3.31`        |
| `CT-DR-B-LP-GMS1-I.3.34-00-1M-20250804-00.docx`         | `CT-DR-B-LP-GMS1-I.3.34-00-1M-20250804-00_CMM.xlsx`    | `CT-DR-B-LP-GMS1-I.3.34-00-1M-20250804-00`            | `yyyy-mm-dd`     | `GMS1 / I.3.34`        |
| `CT-DR-B-LP-GMS2-I.3.34-00-1M-20250806-00.docx`         | `CT-DR-B-LP-GMS2-I.3.34-00-1M-20250806-00_CMM.xlsx`    | `CT-DR-B-LP-GMS2-I.3.34-00-1M-20250806-00`            | `yyyy-mm-dd`     | `GMS2 / I.3.34`        |

---

## Структура репозитория (рекомендация)
```
project-root/
├─ README.md
├─ templates/
│  └─ CommentSheet_Template.xltx
├─ tools/
│  ├─ make_cmm.py
│  └─ watch_folder.py   # опционально
└─ samples/
   ├─ input/   # примеры .docx
   └─ output/  # ожидаемые _CMM.xlsx
```

---

## Примечания
- Если у вас корпоративные политики запрещают `.xltx`, используйте `.xlsx` как шаблон — скрипты поддерживают оба варианта.
- Формат даты `yyyy-mm-dd` выбран для устойчивости (машиночитаемо и однозначно).
- Именованные ячейки позволяют минимизировать зависимость от адресов (если разметка сместится, код не «сломается»).
