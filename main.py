import uvicorn
from database import  engine
from fastapi import FastAPI
from models import *
from routers import router
from urls import task_temp_url

from fastapi.staticfiles import StaticFiles


app = FastAPI(
    title="Logistic",
    version="0.0.1"
)

app.include_router(router)
app.include_router(task_temp_url)
app.mount("/static", StaticFiles(directory='static'), name='static')


@app.get('/')
def index():
    """Тестовый эндпоинт — проверка что сервер работает"""
    return {
        "message": "EcoLogist FastAPI is running",
        "version": "0.0.1",
        "docs": "/docs",
    }


@app.get('/health', include_in_schema=False)
@app.get('/health/')
def health_check():
    """Проверка, что сервер работает"""
    return {"status": "ok", "message": "EcoLogist FastAPI is running"}


if __name__ == '__main__':
    print('Start Server (dev mode)')
    uvicorn.run('main:app', port=8000, host='127.0.0.1', reload=True)
    print('Server Stop')
