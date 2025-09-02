import re
import shutil
from pathlib import Path
from collections import defaultdict
from openpyxl import load_workbook
import tkinter as tk
from tkinter import filedialog

# === НАСТРОЙКИ ===
# Путь к файлу-справочнику
TZ_FILE_PATH = Path("Template/TZ.xlsx")
# Имя листа в справочнике
TZ_SHEET_NAME = "gen_cl"
# Колонки в справочнике
TZ_LOOKUP_COL = "B"  # Колонка с индексами (I.7.5)
TZ_SUFFIX_COL = "J"  # Колонка с суффиксами для папки (KBV)

# Регулярное выражение для извлечения ключа группировки для C-файлов (напр., II.2.6-00-C)
RE_C_GROUPING_KEY = re.compile(
    r"(\b(?:(?:[IVXLCDM]+)\.(?:\d+)(?:\.\d+)?(?:\.\d+)?(?:[A-Za-zА-Яа-я])?)(?:-\d{2}-C))\b",
    re.IGNORECASE
)
# Регулярное выражение для извлечения обычного ключа группировки (напр., I.7.5-00-1G)
RE_GROUPING_KEY = re.compile(
    r"(\b(?:(?:[IVXLCDM]+)\.(?:\d+)(?:\.\d+)?(?:\.\d+)?(?:[A-Za-zА-Яа-я])?)(?:-\d{2}-\w{1,2}))\b",
    re.IGNORECASE
)
# Регулярное выражение для извлечения индекса для поиска (напр., I.7.5)
RE_INDEX_CODE = re.compile(
    r"(\b(?:[IVXLCDM]+)\.(?:\d+)(?:\.\d+)?(?:\.\d+)?(?:[A-Za-zА-Яа-я])?)\b",
    re.IGNORECASE
)

# Карта для транслитерации кириллицы в латиницу для имен папок
CYRILLIC_TO_LATIN = {
    'а': 'a', 'б': 'b', 'в': 'v', 'г': 'g',
    'А': 'A', 'Б': 'B', 'В': 'V', 'Г': 'G'
}

def transliterate_cyrillic_to_latin(text: str) -> str:
    """Конвертирует кириллические символы в латинские по карте."""
    for cyr, lat in CYRILLIC_TO_LATIN.items():
        text = text.replace(cyr, lat)
    return text

def find_suffix_in_tz_file(lookup_key: str) -> str | None:
    """
    Ищет индекс в колонке B листа 'gen_cl' и возвращает значение из колонки J.
    Поддерживает транслитерацию для кириллических букв в конце ключа.
    """
    if not TZ_FILE_PATH.exists():
        print(f"  - [ОШИБКА] Файл-справочник не найден: {TZ_FILE_PATH}")
        return None

    # Карта для транслитерации латиницы в кириллицу для поиска
    translit_map = {'a': 'а', 'b': 'б', 'v': 'в', 'g': 'г'}
    
    keys_to_find = {lookup_key.strip().lower()}
    last_char = lookup_key.strip()[-1]
    
    if last_char in translit_map:
        cyrillic_key = lookup_key.strip()[:-1] + translit_map[last_char]
        keys_to_find.add(cyrillic_key.lower())

    try:
        wb = load_workbook(TZ_FILE_PATH, data_only=True)
        if TZ_SHEET_NAME not in wb.sheetnames:
            print(f"  - [ОШИБКА] Лист '{TZ_SHEET_NAME}' не найден в файле {TZ_FILE_PATH}")
            return None
        
        ws = wb[TZ_SHEET_NAME]
        
        lookup_col_idx = ord(TZ_LOOKUP_COL.upper()) - ord('A') + 1
        suffix_col_idx = ord(TZ_SUFFIX_COL.upper()) - ord('A') + 1

        for row in ws.iter_rows(values_only=True):
            cell_value = str(row[lookup_col_idx - 1]).strip().lower() if row[lookup_col_idx - 1] else ""
            if cell_value in keys_to_find:
                suffix = row[suffix_col_idx - 1]
                return str(suffix).strip() if suffix else None
        
        return None
    except Exception as e:
        print(f"  - [ОШИБКА] Ошибка при чтении файла {TZ_FILE_PATH}: {e}")
        return None

def main():
    """Главная функция для сортировки файлов по папкам."""
    # --- Блок выбора папки ---
    # Для быстрого тестирования: раскомментируйте строку ниже и укажите путь
    source_dir = r"test/toir_tra_index_works"
    
    # # Для обычной работы: закомментируйте строку выше и используйте диалог выбора папки
    # if 'source_dir' not in locals():
    #     root = tk.Tk()
    #     root.withdraw()
    #     print("Пожалуйста, выберите папку с файлами для сортировки...")
    #     source_dir = filedialog.askdirectory(title="Выберите папку для сортировки")

    if not source_dir:
        print("Папка не выбрана. Завершение работы.")
        return
        
    source_path = Path(source_dir)
    print(f"Сканирование директории: {source_path.resolve()}")

    # 1. Группируем все файлы по ключу
    files_by_key = defaultdict(list)
    all_files = [p for p in source_path.rglob('*') if p.is_file()]

    for file_path in all_files:
        # Сначала ищем более специфичный ключ для C-файлов
        c_match = RE_C_GROUPING_KEY.search(file_path.name)
        if c_match:
            key = c_match.group(1)
            files_by_key[key].append(file_path)
            continue  # Файл уже сгруппирован, переходим к следующему

        # Если не C-файл, ищем обычный ключ
        match = RE_GROUPING_KEY.search(file_path.name)
        if match:
            key = match.group(1)
            files_by_key[key].append(file_path)

    if not files_by_key:
        print("Не найдено файлов, соответствующих шаблону имен для сортировки.")
        return

    print(f"Найдено {len(files_by_key)} групп файлов для сортировки.")

    # 2. Обрабатываем каждую группу
    for key, file_paths in files_by_key.items():
        print(f"\n--- Обработка группы: {key} ---")

        folder_name: str
        # Проверяем, является ли ключ C-ключом
        if key.upper().endswith("-C"):
            folder_name = transliterate_cyrillic_to_latin(key)
            print(f"  - Группа C-файлов. Имя папки: {folder_name}")
        else:
            # Стандартная логика с поиском суффикса
            index_match = RE_INDEX_CODE.search(key)
            if not index_match:
                print(f"  - [ПРЕДУПРЕЖДЕНИЕ] Не удалось извлечь индекс из ключа '{key}'. Пропуск группы.")
                continue

            index_code = index_match.group(1)
            print(f"  - Поиск суффикса для индекса: {index_code}")

            suffix = find_suffix_in_tz_file(index_code)
            if not suffix:
                print(f"  - [ПРЕДУПРЕПРЕЖДЕНИЕ] Суффикс для индекса '{index_code}' не найден в {TZ_FILE_PATH}. Пропуск группы.")
                continue

            print(f"  - Найден суффикс: '{suffix}'")
            latin_key = transliterate_cyrillic_to_latin(key)
            folder_name = f"{latin_key}_{suffix}"

        # Создаем папку и перемещаем файлы
        dest_dir = source_path / folder_name
        
        print(f"  - Создание папки: {dest_dir.name}")
        dest_dir.mkdir(exist_ok=True)

        # 4. Перемещаем файлы
        for file_path in file_paths:
            dest_file = dest_dir / file_path.name
            try:
                print(f"    - Перемещение: {file_path.name}")
                shutil.move(str(file_path), str(dest_file))
            except Exception as e:
                print(f"    - [ОШИБКА] Не удалось переместить {file_path.name}: {e}")

    print("\nОбработка завершена.")

if __name__ == "__main__":
    main()