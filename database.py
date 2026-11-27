from sqlalchemy.engine import create_engine
from sqlalchemy.orm import DeclarativeBase


engine = create_engine(
    "sqlite:///db.sqlite",
    echo=True
)



class Model(DeclarativeBase):
    pass


# Импортируем все модели, чтобы они были зарегистрированы
from models import TaskModel, UserModel, AddressModel  # noqa: F401


# Создаём все таблицы
Model.metadata.create_all(engine)
