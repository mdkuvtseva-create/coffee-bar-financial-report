"""
Скрипт для объединения всех данных в один Excel файл
Каждый файл будет на отдельном листе (вкладке)
"""

import pandas as pd
from pathlib import Path

print("=" * 70)
print("📊 СОЗДАНИЕ ЕДИНОГО EXCEL ФАЙЛА С НЕСКОЛЬКИМИ ЛИСТАМИ")
print("=" * 70)

# ============================================================
# ПУТИ К ФАЙЛАМ
# ============================================================
DATA_DIR = Path("C:/Users/Диана/Desktop/Проекты/processed")
OUTPUT_FILE = Path("C:/Users/Диана/Desktop/Проекты/Финансовый_отчет_2022_2023.xlsx")

print(f"📂 Папка с обработанными файлами: {DATA_DIR}")
print(f"📄 Выходной файл: {OUTPUT_FILE}")
print("")

# ============================================================
# ПРОВЕРКА НАЛИЧИЯ ФАЙЛОВ
# ============================================================
print("🔍 ПРОВЕРКА ФАЙЛОВ:")
print("-" * 50)

files_needed = [
    'main_checking_2022_2023.csv',
    'secondary_checking_2022_2023.csv',
    'credit_card_2022_2023.csv',
    'payroll_2022_2023.csv'
]

all_found = True
for filename in files_needed:
    file_path = DATA_DIR / filename
    if file_path.exists():
        print(f"  ✅ {filename} - НАЙДЕН")
    else:
        print(f"  ❌ {filename} - НЕ НАЙДЕН!")
        all_found = False

if not all_found:
    print("\n⚠️ НЕ ВСЕ ФАЙЛЫ НАЙДЕНЫ!")
    print("   Сначала запустите скрипт clean_data.py для обработки данных.")
    exit()

print("\n✅ Все файлы найдены! Начинаем создание Excel...")
print("=" * 70 + "\n")

# ============================================================
# СОЗДАНИЕ EXCEL ФАЙЛА
# ============================================================

# Создаем объект для записи Excel
with pd.ExcelWriter(OUTPUT_FILE, engine='openpyxl') as writer:
    
    # 1. Основной счет
    print("📊 Загрузка основного счета...")
    df_main = pd.read_csv(DATA_DIR / 'main_checking_2022_2023.csv')
    df_main.to_excel(writer, sheet_name='Основной счет', index=False)
    print(f"   ✅ Загружено {len(df_main)} строк")
    
    # 2. Второй счет
    print("📊 Загрузка второго счета...")
    df_sec = pd.read_csv(DATA_DIR / 'secondary_checking_2022_2023.csv')
    df_sec.to_excel(writer, sheet_name='Второй счет', index=False)
    print(f"   ✅ Загружено {len(df_sec)} строк")
    
    # 3. Кредитная карта
    print("📊 Загрузка кредитной карты...")
    df_cc = pd.read_csv(DATA_DIR / 'credit_card_2022_2023.csv')
    df_cc.to_excel(writer, sheet_name='Кредитная карта', index=False)
    print(f"   ✅ Загружено {len(df_cc)} строк")
    
    # 4. Зарплаты
    print("📊 Загрузка зарплат...")
    df_pay = pd.read_csv(DATA_DIR / 'payroll_2022_2023.csv')
    df_pay.to_excel(writer, sheet_name='Зарплаты', index=False)
    print(f"   ✅ Загружено {len(df_pay)} строк")
    
    # 5. Сводная статистика (отдельный лист)
    print("📊 Создание листа со статистикой...")
    
    # Читаем статистику из текстового файла
    stats_file = DATA_DIR / 'summary_statistics.txt'
    if stats_file.exists():
        with open(stats_file, 'r', encoding='utf-8') as f:
            stats_text = f.read()
        
        # Создаем DataFrame для статистики
        stats_lines = stats_text.strip().split('\n')
        stats_data = []
        for line in stats_lines:
            if ':' in line and not line.startswith('='):
                parts = line.split(':', 1)
                if len(parts) == 2:
                    key = parts[0].strip()
                    value = parts[1].strip()
                    stats_data.append([key, value])
        
        df_stats = pd.DataFrame(stats_data, columns=['Показатель', 'Значение'])
        df_stats.to_excel(writer, sheet_name='Статистика', index=False)
        print(f"   ✅ Статистика добавлена")
    else:
        print("   ⚠️ Файл со статистикой не найден")

print("\n" + "=" * 70)
print("✅ EXCEL ФАЙЛ СОЗДАН!")
print("=" * 70)
print(f"\n📄 Файл сохранен по адресу:")
print(f"   {OUTPUT_FILE}")

# Показываем размер файла
if OUTPUT_FILE.exists():
    size_kb = OUTPUT_FILE.stat().st_size / 1024
    size_mb = size_kb / 1024
    if size_mb > 1:
        print(f"   Размер: {size_mb:.1f} MB")
    else:
        print(f"   Размер: {size_kb:.1f} KB")

print("\n📂 Листы в файле:")
print("   📄 Лист 1: Основной счет")
print("   📄 Лист 2: Второй счет")
print("   📄 Лист 3: Кредитная карта")
print("   📄 Лист 4: Зарплаты")
print("   📄 Лист 5: Статистика")

print("\n" + "=" * 70)
print("🎉 ГОТОВО! Откройте файл в Excel.")
print("=" * 70)
