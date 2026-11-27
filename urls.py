from fastapi import File, Request, APIRouter, UploadFile, HTTPException
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from fastapi.responses import RedirectResponse, JSONResponse
from datetime import datetime
import os

from openpyxl import load_workbook, Workbook
from fanc import build_address_lookup, create_report, create_report_from_db, get_all_addresses, get_unique_drivers, update_driver, process_sell_files


task_temp_url = APIRouter(tags=['/api/tasks'])
template = Jinja2Templates(directory='templates')

# Настройки обработки
DRIVE_FILE = 'download/BaseDrive.xlsx'
DRIVE_SHEET = None
DRIVE_COLUMNS = {
    'address': 'A',
    'buyer': 'B',
    'driver': 'C',
    'district': 'D',
}
SELL_COLUMNS = {
    'address': 'L',
    'buyer': 'H',
    'comment': 'K',
    'division': 'M',
    'date': 'N',
    'status': 'U',
    'maney': 'I',
}
SELL_SHEET = None
SKIP_EMPTY = True

# Хранение данных для редактирования (в памяти)
pending_reports = {}


@task_temp_url.get(path='/')
def index(request: Request):
    return template.TemplateResponse(
        request=request,
        name='home.html',
    )


@task_temp_url.get(path='/tasklist/')
def add_task(request: Request):
    return template.TemplateResponse(
        request=request,
        name='tasklist.html',
    )


@task_temp_url.get(path='/tasklist/create/')
def add_list(request: Request):
    return template.TemplateResponse(
        request=request,
        name='create_list.html',
    )


@task_temp_url.get("/tasklist/addresses/")
async def list_addresses(request: Request):
    """Отображает таблицу всех адресов для редактирования"""
    addresses = get_all_addresses()
    return template.TemplateResponse(
        request=request,
        name='addresses.html',
        context={"addresses": addresses}
    )


@task_temp_url.put("/tasklist/update-driver/")
async def update_driver_endpoint(request: Request):
    """Обновляет водителя в БД"""
    data = await request.json()
    address_id = data.get('id')
    driver = data.get('driver', '')
    
    if not address_id:
        raise HTTPException(status_code=400, detail="ID адреса не указан")
    
    success = update_driver(address_id, driver)
    
    if success:
        return JSONResponse(content={"message": "Водитель обновлён"})
    else:
        raise HTTPException(status_code=404, detail="Адрес не найден")


@task_temp_url.post("/upload")
async def upload_files(
    request: Request,
    file1: UploadFile = File(...),
    file2: UploadFile = File(...),
):
    upload_dir = "download"
    os.makedirs(upload_dir, exist_ok=True)

    # 1. Сохраняем загруженные файлы
    saved_files = []
    for file in [file1, file2]:
        filename = file.filename
        if filename is not None:
            file_path = os.path.join(upload_dir, filename)
            content = await file.read()
            with open(file_path, "wb") as buffer:
                buffer.write(content)
            saved_files.append(file_path)

    # 2. Обработка файлов
    if len(saved_files) == 2:
        # Фильтр по столбцу M (division)
        DIVISION_FILTER = [
            'Склад и логистика',
            'Склад Готовой Продукции (Производство СоусПак Красноярск)',
        ]
        
        # Обрабатываем файлы и получаем данные для таблицы
        result = process_sell_files(
            sell_files=saved_files,
            sell_sheet_name=SELL_SHEET,
            cols=SELL_COLUMNS,
            skip_empty=SKIP_EMPTY,
            verbose=True,
            division_filter=DIVISION_FILTER,
        )

        # Генерируем ключ для хранения данных
        report_key = f"report_{datetime.now().strftime('%Y%m%d_%H%M%S')}"
        pending_reports[report_key] = {
            'saved_files': saved_files,
            'rows': result['rows'],
            'stats': result['stats'],
            'district_sums': result['district_sums'],
            'report_filename': f"{datetime.now().strftime('%d.%m.%Y')}.xlsx",
        }

        drivers = get_unique_drivers()
        
        context = {
            "request": request,
            "file1_name": file1.filename or "",
            "file2_name": file2.filename or "",
            "report_key": report_key,
            "rows": result['rows'],
            "found": result['stats']['found'],
            "not_found": result['stats']['not_found'],
            "total_sum": sum(result['district_sums'].values()),
            "district_sums": result['district_sums'],
            "drivers": drivers,
        }
    else:
        context = {
            "request": request,
            "error": "Загрузите ровно два файла.",
        }

    return template.TemplateResponse(
        request=request,
        name='report_edit.html',
        context=context,
    )


@task_temp_url.post("/generate-report/")
async def generate_report(request: Request):
    """Генерирует Excel-отчёт на основе отредактированных данных"""
    data = await request.json()
    report_key = data.get('report_key')
    
    if not report_key or report_key not in pending_reports:
        raise HTTPException(status_code=404, detail="Данные не найдены")
    
    report_data = pending_reports[report_key]
    
    # Создаём папку для отчётов
    report_dir = "static/upload"
    os.makedirs(report_dir, exist_ok=True)
    
    report_path = os.path.join(report_dir, report_data['report_filename'])
    
    # Генерируем отчёт
    stats = create_report_from_db(
        sell_files=report_data['saved_files'],
        report_file=report_path,
        sell_sheet_name=SELL_SHEET,
        cols=SELL_COLUMNS,
        skip_empty=SKIP_EMPTY,
        verbose=True,
    )
    
    # Удаляем из pending
    del pending_reports[report_key]
    
    return JSONResponse(content={
        "message": "Отчёт сгенерирован",
        "report_filename": report_data['report_filename'],
        "found": stats['found'],
        "total_sum": sum(stats['district_sums'].values()),
    })


@task_temp_url.get("/tasklist/report/{report_filename}")
async def view_report(request: Request, report_filename: str):
    """Отображает финальный отчёт"""
    report_path = os.path.join("static/upload", report_filename)
    
    if not os.path.exists(report_path):
        return template.TemplateResponse(
            request=request,
            name='report.html',
            context={"request": request, "error": "Отчёт не найден"}
        )
    
    # Считаем сумму из Excel
    wb = load_workbook(report_path)
    ws = wb.active
    if ws is None:
        return template.TemplateResponse(
            request=request,
            name='report.html',
            context={"request": request, "error": "Не удалось открыть отчёт"}
        )
    total_sum = 0.0
    found = 0
    for row in ws.iter_rows(min_row=2):
        val = row[4].value
        if val is not None and isinstance(val, (int, float)):
            total_sum += val
            found += 1
    wb.close()
    
    return template.TemplateResponse(
        request=request,
        name='report.html',
        context={
            "request": request,
            "found": found,
            "total_sum": total_sum,
            "report_filename": report_filename,
        }
    )
