from openpyxl import load_workbook, Workbook
from datetime import datetime
import os


def col_to_index(col_letter: str) -> int:
    """Преобразует букву столбца ('A', 'B', 'AA') в номер (1-based)."""
    col_letter = col_letter.upper().strip()
    result = 0
    for char in col_letter:
        result = result * 26 + (ord(char) - ord('A') + 1)
    return result


def build_address_lookup(drive_file: str, drive_sheet_name: str | None,
                         cols: dict) -> dict[str, dict]:
    """
    Читает файл с адресами и строит словарь:
    { 'адрес': { 'buyer': ..., 'driver': ..., 'district': ... } }
    """
    wb = load_workbook(drive_file, read_only=True)
    ws = wb[drive_sheet_name] if drive_sheet_name else wb.active

    if ws is None:
        raise ValueError(f"Лист '{drive_sheet_name}' не найден в файле {drive_file}")

    address_idx = col_to_index(cols['address'])
    buyer_idx = col_to_index(cols['buyer'])
    driver_idx = col_to_index(cols['driver'])
    district_idx = col_to_index(cols['district'])

    lookup: dict[str, dict] = {}

    for row in ws.iter_rows(min_row=2):  # пропускаем заголовок
        address_raw = row[address_idx - 1].value
        if address_raw is None:
            continue

        address_key = str(address_raw).strip().lower()

        if address_key in lookup:
            continue

        lookup[address_key] = {
            'buyer':    str(row[buyer_idx - 1].value).strip() if row[buyer_idx - 1].value else '',
            'driver':   str(row[driver_idx - 1].value).strip() if row[driver_idx - 1].value else '',
            'district': str(row[district_idx - 1].value).strip() if row[district_idx - 1].value else '',
        }

    wb.close()
    return lookup


def create_report(sell_files: list[str], report_file: str, lookup: dict[str, dict],
                  sell_sheet_name: str | None, cols: dict,
                  skip_empty: bool, verbose: bool = True) -> dict:
    """
    Создаёт отчётный файл из нескольких файлов 1С.
    Лист "Отчёт": № | Адрес | Покупатель | Район | Сумма
    Лист "Суммы по районам": Район | Сумма

    Возвращает статистику.
    """
    sell_addr_letter = cols['address']
    sell_buyer_letter = cols['buyer']
    sell_maney_letter = cols.get('maney')

    stats = {
        'total': 0, 'found': 0, 'not_found': 0, 'empty': 0,
        'district_sums': {}, 'files': {}
    }

    # Создаём новый workbook для отчёта
    wb_report = Workbook()
    ws_report = wb_report.active
    if ws_report is None:
        raise ValueError("Не удалось создать лист отчёта")
    ws_report.title = 'Отчёт'

    # Заголовки
    headers = ['№', 'Адрес', 'Покупатель', 'Район', 'Сумма']
    for col_idx, header in enumerate(headers, start=1):
        ws_report.cell(row=1, column=col_idx, value=header)

    # Заполняем данные из всех файлов
    row_num = 2

    for sell_file in sell_files:
        file_stats = {'total': 0, 'found': 0, 'not_found': 0, 'empty': 0}

        if verbose:
            print(f"  Обработка: {sell_file}")

        wb = load_workbook(sell_file)
        ws = wb[sell_sheet_name] if sell_sheet_name else wb.active

        if ws is None:
            raise ValueError(f"Лист '{sell_sheet_name}' не найден в файле {sell_file}")

        for row_num_source in range(2, ws.max_row + 1):
            addr_raw = ws[f'{sell_addr_letter}{row_num_source}'].value
            file_stats['total'] += 1

            if addr_raw is None or str(addr_raw).strip() == '':
                file_stats['empty'] += 1
                if skip_empty:
                    continue

            addr_key = str(addr_raw).strip().lower()
            buyer_raw = ws[f'{sell_buyer_letter}{row_num_source}'].value

            maney_value = 0.0
            if sell_maney_letter:
                maney_cell = ws[f'{sell_maney_letter}{row_num_source}']
                maney_raw = maney_cell.value
                try:
                    maney_value = float(maney_raw) if maney_raw is not None else 0.0
                except (ValueError, TypeError):
                    maney_value = 0.0

            if addr_key in lookup:
                info = lookup[addr_key]
                district = info['district'] if info['district'] else 'Без района'
                file_stats['found'] += 1
            else:
                district = 'Не найден'
                file_stats['not_found'] += 1
                maney_value = 0.0  # не найденные адреса не учитываем

            # Суммируем по районам (только для найденных адресов)
            if district != 'Не найден':
                if district not in stats['district_sums']:
                    stats['district_sums'][district] = 0.0
                stats['district_sums'][district] += maney_value

            ws_report.cell(row=row_num, column=1, value=row_num - 1)
            ws_report.cell(row=row_num, column=2, value=addr_raw)
            ws_report.cell(row=row_num, column=3, value=buyer_raw)
            ws_report.cell(row=row_num, column=4, value=district)
            ws_report.cell(row=row_num, column=5, value=maney_value)
            row_num += 1

        wb.close()

        # Сохраняем статистику по файлу
        stats['files'][sell_file] = file_stats
        stats['total'] += file_stats['total']
        stats['found'] += file_stats['found']
        stats['not_found'] += file_stats['not_found']
        stats['empty'] += file_stats['empty']

    # Создаём лист "Суммы по районам"
    total_all = 0.0
    last_row = 1
    if stats['district_sums']:
        ws_district = wb_report.create_sheet('Суммы по районам')
        ws_district.cell(row=1, column=1, value='Район')
        ws_district.cell(row=1, column=2, value='Сумма')

        for row_idx, (district, amount) in enumerate(
            sorted(stats['district_sums'].items()), start=2
        ):
            ws_district.cell(row=row_idx, column=1, value=district)
            ws_district.cell(row=row_idx, column=2, value=amount)
            total_all += amount
            last_row = row_idx

        ws_district.cell(row=last_row + 1, column=1, value='ИТОГО')
        ws_district.cell(row=last_row + 1, column=2, value=total_all)

    wb_report.save(report_file)
    wb_report.close()

    return stats


# =============================================================================
# НАСТРОЙКИ — для запуска как standalone скрипта
# =============================================================================

SELL_FILES = [
    'download/Eco.xlsx',
    'download/Sous.xlsx',
]

DRIVE_FILE = 'download/BaseDrive.xlsx'

REPORT_FILE = 'static/upload/{time}.xlsx'

SELL_SHEET = None

SELL_COLUMNS = {
    'address':   'L',
    'buyer':     'H',
    'comment':   'K',
    'division':  'M',
    'date':      'N',
    'status':    'U',
    'maney':     'I',
}

DRIVE_SHEET = None
DRIVE_COLUMNS = {
    'address':   'A',
    'buyer':     'B',
    'driver':    'C',
    'district':  'D',
}

SKIP_EMPTY_ADDRESS = True
VERBOSE = True


def main():
    verbose = VERBOSE

    if verbose:
        print("=" * 60)
        print("Скрипт создания отчёта по отгрузкам")
        print("=" * 60)
        print(f"\nФайлы 1С:")
        for f in SELL_FILES:
            print(f"  - {f}")
        print(f"Файл адресов:  {DRIVE_FILE}")
        print(f"Отчёт:         {REPORT_FILE}")
        print()

    # 1. Строим справочник
    if verbose:
        print("Чтение файла с адресами контрагентов...")
    lookup = build_address_lookup(
        drive_file=DRIVE_FILE,
        drive_sheet_name=DRIVE_SHEET,
        cols=DRIVE_COLUMNS
    )
    if verbose:
        print(f"  Найдено адресов: {len(lookup)}")
    print()

    # 2. Создаём отчёт
    report = REPORT_FILE
    if '{time}' in report:
        report = report.replace('{time}', datetime.now().strftime('%d.%m.%Y'))
    report_dir = os.path.dirname(report)
    if report_dir:
        os.makedirs(report_dir, exist_ok=True)

    if verbose:
        print("Создание отчёта...")
    stats = create_report(
        sell_files=SELL_FILES,
        report_file=report,
        lookup=lookup,
        sell_sheet_name=SELL_SHEET,
        cols=SELL_COLUMNS,
        skip_empty=SKIP_EMPTY_ADDRESS,
        verbose=VERBOSE
    )

    # 3. Выводим результат
    print()
    print("=" * 60)
    print("РЕЗУЛЬТАТ")
    print("=" * 60)

    for fname, fstats in stats['files'].items():
        print(f"\n  {os.path.basename(fname)}:")
        print(f"    Всего:       {fstats['total']}")
        print(f"    Найдены:     {fstats['found']}")
        print(f"    Не найдены:  {fstats['not_found']}")
        print(f"    Пустые:      {fstats['empty']}")

    print()
    print(f"  ИТОГО строк:       {stats['total']}")
    print(f"  ИТОГО найдены:     {stats['found']}")
    print(f"  ИТОГО не найдены:  {stats['not_found']}")
    print(f"  ИТОГО пустые:      {stats['empty']}")

    # Суммы по районам
    if stats['district_sums']:
        print()
        print("=" * 60)
        print("СУММЫ ПО РАЙОНАМ")
        print("=" * 60)
        total_all = 0.0
        for district, amount in sorted(stats['district_sums'].items()):
            print(f"  {district}: {amount:,.2f}")
            total_all += amount
        print("-" * 40)
        print(f"  ИТОГО: {total_all:,.2f}")
        print("=" * 60)

    print(f"\nОтчёт сохранён: {report}")
    print("=" * 60)


if __name__ == '__main__':
    main()
