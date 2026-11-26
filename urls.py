from fastapi import File, Request, APIRouter, UploadFile
from fastapi.templating import Jinja2Templates
from fastapi.staticfiles import StaticFiles
from fastapi.responses import RedirectResponse
from datetime import datetime
import os

from fanc import build_address_lookup, create_report


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


@task_temp_url.post("/upload")
async def upload_files(
    request: Request,
    file1: UploadFile = File(...),
    file2: UploadFile = File(...),
):
    upload_dir = "download"
    os.makedirs(upload_dir, exist_ok=True)

    # Папка для отчётов
    report_dir = "static/upload"
    os.makedirs(report_dir, exist_ok=True)

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

    # 2. Запускаем обработку через fanc.py
    if len(saved_files) == 2:
        # Строим справочник адресов
        lookup = build_address_lookup(
            drive_file=DRIVE_FILE,
            drive_sheet_name=DRIVE_SHEET,
            cols=DRIVE_COLUMNS,
        )

        # Создаём отчёт
        report_filename = f"{datetime.now().strftime('%d.%m.%Y')}.xlsx"
        report_path = os.path.join(report_dir, report_filename)

        stats = create_report(
            sell_files=saved_files,
            report_file=report_path,
            lookup=lookup,
            sell_sheet_name=SELL_SHEET,
            cols=SELL_COLUMNS,
            skip_empty=SKIP_EMPTY,
            verbose=True,
        )

        # Считаем общую сумму (только по найденным адресам)
        total_sum = sum(
            amount for district, amount in stats['district_sums'].items()
            if district != 'Не найден'
        )

        context = {
            "request": request,
            "file1_name": file1.filename or "",
            "file2_name": file2.filename or "",
            "found": stats['found'],
            "not_found": stats['not_found'],
            "empty": stats['empty'],
            "total": stats['total'],
            "total_sum": total_sum,
            "report_filename": report_filename,
            "district_sums": stats['district_sums'],
        }
    else:
        context = {
            "request": request,
            "error": "Загрузите ровно два файла.",
        }

    return template.TemplateResponse(
        request=request,
        name='report.html',
        context=context,
    )
