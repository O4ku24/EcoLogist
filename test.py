"""
Скрипт для импорта данных из BaseDrive.xlsx в БД.
Запустите один раз для начального импорта адресов.
"""
from fanc import xlsx_to_csv, import_csv_to_db

DRIVE_FILE = 'download/BaseDrive.xlsx'
CSV_FILE = 'download/BaseDrive.csv'

print("=" * 60)
print("IMPORT адресов из BaseDrive.xlsx in DB")
print("=" * 60)

# 1. Конвертируем XLSX в CSV
print(f"\n1. Convert {DRIVE_FILE} -> {CSV_FILE}...")
try:
    xlsx_to_csv(DRIVE_FILE, CSV_FILE, {
        'address': 'A',
        'buyer': 'B',
        'driver': 'C',
        'district': 'D',
    })
    print("   OK: Conversion completed")
except Exception as e:
    print(f"   ERROR: {e}")
    exit(1)

# 2. Импортируем CSV в БД
print(f"\n2. Import {CSV_FILE} to DB...")
try:
    count = import_csv_to_db(CSV_FILE)
    print(f"   OK: Imported {count} records")
except Exception as e:
    print(f"   ERROR: {e}")
    exit(1)

print("\n" + "=" * 60)
print("IMPORT COMPLETED")
print("=" * 60)