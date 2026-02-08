"""
SQLAlchemy 数据库初始化
"""
from sqlalchemy import create_engine
from sqlalchemy.orm import declarative_base, sessionmaker

from app.core.config import DATABASE_URL, SQLITE_CONNECT_ARGS

engine = create_engine(
    DATABASE_URL,
    connect_args=SQLITE_CONNECT_ARGS
)

SessionLocal = sessionmaker(autocommit=False, autoflush=False, bind=engine, expire_on_commit=False)
Base = declarative_base()


def init_db():
    from app.db import models  # noqa: F401
    Base.metadata.create_all(bind=engine)
