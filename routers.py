from fastapi import APIRouter, Request, HTTPException
from fastapi.responses import JSONResponse
from schemas import TaskCreateSchema, TaskUpdateSchema, TaskSchema, UserAddSchema, UserCreateSchema, UserUpdateSchema
from models import TaskModel, UserModel, RangUser
from sqlalchemy.orm import Session
from sqlalchemy import select, insert, update, delete
from database import engine
from datetime import datetime

router = APIRouter(prefix='/api/tasks')


@router.get('/health/')
def health_check():
    """Проверка, что сервер работает"""
    return {"status": "ok", "message": "EcoLogist FastAPI is running"}


@router.get('/create/')
def get_list_task(request: Request):
    """Получить список всех задач"""
    with Session(engine) as session:
        stmt = select(TaskModel)
        tasks = session.execute(stmt).scalars().all()
        return [
            {
                "task_id": t.task_id,
                "user_id": t.user_id,
                "title": t.title,
                "description": t.description,
                "status": t.status,
                "create_at": t.create_at.isoformat() if t.create_at else None,
            }
            for t in tasks
        ]


@router.post('/create/')
def add_task(request: Request, task: TaskCreateSchema):
    """Создать новую задачу"""
    with Session(engine) as session:
        stmt = select(TaskModel).order_by(TaskModel.task_id).descending().limit(1)
        last = session.execute(stmt).scalar_one_or_none()
        new_id = (last.task_id + 1) if last else 1

        new_task = TaskModel(
            user_id=task.user_id if hasattr(task, 'user_id') else 1,
            task_id=new_id,
            title=task.title,
            description=task.description,
            status=False,
        )
        session.add(new_task)
        session.commit()
        session.refresh(new_task)

        return {
            "task_id": new_task.task_id,
            "title": new_task.title,
            "description": new_task.description,
            "status": new_task.status,
            "create_at": new_task.create_at.isoformat() if new_task.create_at else None,
        }


@router.put('/list_task/')
def update_task(request: Request, task_id: int, task_chenge: TaskUpdateSchema):
    """Обновить задачу"""
    with Session(engine) as session:
        task = session.get(TaskModel, task_id)
        if not task:
            raise HTTPException(status_code=404, detail="Задача не найдена")

        task.title = task_chenge.title
        task.description = task_chenge.description
        task.status = task_chenge.status
        session.commit()

        return {
            "task_id": task.task_id,
            "title": task.title,
            "description": task.description,
            "status": task.status,
        }


@router.delete('/delete/')
def delete_task(request: Request, task_id: int):
    """Удалить задачу"""
    with Session(engine) as session:
        task = session.get(TaskModel, task_id)
        if not task:
            raise HTTPException(status_code=404, detail="Задача не найдена")

        session.delete(task)
        session.commit()

        return {"message": f"Задача {task_id} удалена"}