import re
import shutil
from pathlib import Path
import sys

# Настройка UTF-8 вывода в Windows-консоли
try:
    sys.stdout.reconfigure(encoding="utf-8", errors="replace")
    sys.stderr.reconfigure(encoding="utf-8", errors="replace")
except Exception:
    pass

# === НАСТРОЙКИ ПУТЕЙ ===
# Замените пути на актуальные для вашей системы.
# Используйте r"..." для путей в Windows, чтобы избежать проблем с \

# 1. Папка, где лежат исходные проекты для обработки (внутри которой папки с отчетами)
INBOX_DIR = Path(r"test/toir_raspredelenije/00_Inbox")

# 2. Папка для копирования PDF-файлов (примечания)
NOTES_DIR = Path(r"test/toir_raspredelenije/03_Notes")

# 3. Папка для копирования PDF-файлов (трансмиттал)
TRA_GST_DIR = Path(r"test/toir_raspredelenije/04_TRA_GST")

# 4. Корневая папка для создаваемой структуры (Год/Месяц/Part/pdf и Native)
DEST_ROOT_DIR = Path(r"test/toir_raspredelenije")

# 5. Временная папка для создания архивов (может совпадать с DEST_ROOT_DIR)
TEMP_ARCHIVE_DIR = Path(r"test/toir_raspredelenije")

# === НАСТРОЙКИ ЛОГИКИ ===

# Регулярное выражение для разбора имени файла
# Извлекает part (LP или CS), object_name и date (ггггммдд)
RE_FILENAME = re.compile(
    r"^CT-(?:DR|CL)-B-(?P<part>LP|CS)-(?P<object_name>[A-Z0-9]+)-.*?(?P<date>\d{8})-\d{2}_All\.pdf$",
    re.IGNORECASE
)

# Словарь для перевода номера месяца в название
MONTH_MAP = {
    "01": "January", "02": "February", "03": "March", "04": "April",
    "05": "May", "06": "June", "07": "July", "08": "August",
    "09": "September", "10": "October", "11": "November", "12": "December"
}

def normalize_object_name(object_name: str) -> str:
    """
    Нормализует имя объекта, удаляя ведущий ноль для однозначных номеров.
    Пример: BVS05 -> BVS5. BVS10 -> BVS10.
    """
    # Ищем шаблон: BVS, затем 0, затем одна цифра от 1 до 9
    match = re.match(r"^(BVS)0([1-9])$", object_name, re.IGNORECASE)
    if match:
        # Собираем новое имя из первой группы (BVS) и второй (цифра)
        normalized_name = match.group(1) + match.group(2)
        print(f"  - [ИНФО] Имя объекта нормализовано: {object_name} -> {normalized_name}")
        return normalized_name
    return object_name

def process_project_folder(project_path: Path):
    """Обрабатывает одну папку из INBOX_DIR."""
    print(f"\n--- Сканирование папки: {project_path.name} ---")

    # 1. Ищем уникальный *_All.pdf файл
    all_pdf_files = list(project_path.glob("*_All.pdf"))
    if not all_pdf_files:
        print("  - [ИНФО] Файл *_All.pdf не найден. Пропуск папки.")
        return
    if len(all_pdf_files) > 1:
        print(f"  - [ПРЕДУПРЕЖДЕНИЕ] Найдено несколько ({len(all_pdf_files)}) *_All.pdf файлов. Пропуск папки.")
        return
    
    pdf_file = all_pdf_files[0]
    print(f"  - Найден файл для обработки: {pdf_file.name}")

    # 2. Разбираем имя файла
    match = RE_FILENAME.match(pdf_file.name)
    if not match:
        print(f"  - [ПРЕДУПРЕЖДЕНИЕ] Имя файла {pdf_file.name} не соответствует шаблону. Пропуск.")
        return

    data = match.groupdict()
    date_str = data["date"]
    year = date_str[:4]
    month_num = date_str[4:6]
    month_name = MONTH_MAP.get(month_num, "UnknownMonth")
    part = data["part"].upper()
    
    # Получаем и нормализуем имя объекта
    object_name_raw = data["object_name"].upper()
    object_name = normalize_object_name(object_name_raw)
    
    print(f"  - Разобраны данные: Год={year}, Месяц={month_num}.{month_name}, Part={part}, Объект={object_name} (исходное: {object_name_raw})")

    # 3. Копируем PDF в простые директории
    try:
        print(f"  - Копирование в {NOTES_DIR.name}...")
        shutil.copy(pdf_file, NOTES_DIR)
        print(f"  - Копирование в {TRA_GST_DIR.name}...")
        shutil.copy(pdf_file, TRA_GST_DIR)
    except Exception as e:
        print(f"  - [ОШИБКА] Не удалось скопировать PDF: {e}")
        return

    # 4. Копируем PDF по сложному пути
    try:
        month_folder_name = f"{month_num}.{month_name}"
        pdf_dest_dir = DEST_ROOT_DIR / year / month_folder_name / part / "pdf" / object_name
        pdf_dest_dir.mkdir(parents=True, exist_ok=True)
        print(f"  - Копирование PDF в: {pdf_dest_dir}")
        shutil.copy(pdf_file, pdf_dest_dir)
    except Exception as e:
        print(f"  - [ОШИБКА] Не удалось скопировать PDF по сложному пути: {e}")

    # 5. Архивируем исходную папку и перемещаем архив
    try:
        # Имя архива будет как имя папки, в которой он лежит
        archive_basename = TEMP_ARCHIVE_DIR / project_path.name
        print(f"  - Создание архива для папки: {project_path.name}...")
        archive_path_str = shutil.make_archive(str(archive_basename), 'zip', str(project_path))
        archive_path = Path(archive_path_str)
        
        month_folder_name = f"{month_num}.{month_name}"
        archive_dest_dir = DEST_ROOT_DIR / year / month_folder_name / part / "Native" / object_name
        archive_dest_dir.mkdir(parents=True, exist_ok=True)
        
        print(f"  - Перемещение архива в: {archive_dest_dir}")
        shutil.move(str(archive_path), archive_dest_dir)
    except Exception as e:
        print(f"  - [ОШИБКА] Не удалось создать или переместить архив: {e}")

def main():
    """Главная функция для распределения файлов."""
    print("Запуск скрипта по распределению файлов...")

    # Проверяем и создаем базовые папки, если они не существуют
    # Это важно для папок назначения, чтобы shutil.copy не выдавал ошибку
    for dir_path in [NOTES_DIR, TRA_GST_DIR]:
        if not dir_path.exists():
            print(f"Создание необходимой папки: {dir_path}")
            dir_path.mkdir(parents=True, exist_ok=True)

    if not INBOX_DIR.exists():
        print(f"[КРИТИЧЕСКАЯ ОШИБКА] Папка с исходными файлами не найдена: {INBOX_DIR}")
        return

    # Ищем папки в INBOX_DIR
    project_folders = [p for p in INBOX_DIR.iterdir() if p.is_dir()]

    if not project_folders:
        print(f"В папке {INBOX_DIR} не найдено подпапок для обработки.")
        return

    print(f"Найдено {len(project_folders)} папок для проверки.")
    for folder in project_folders:
        process_project_folder(folder)

    print("\nОбработка завершена.")

if __name__ == "__main__":
    main()