"""
SQLAlchemy 数据库初始化
"""
from sqlalchemy import create_engine, inspect, text
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
    ensure_model_config_columns()
    ensure_parse_record_columns()


def ensure_model_config_columns():
    """为已存在的 model_config 表补齐新增列（create_all 不会给旧表加列）。
    SQLite / MySQL 均支持 ALTER TABLE ADD COLUMN，新增列均为 nullable。"""
    try:
        insp = inspect(engine)
        if not insp.has_table("model_config"):
            return
        existing = {c["name"] for c in insp.get_columns("model_config")}
    except Exception:
        return
    additions = [
        ("temperature", "FLOAT"),
        ("model_mode", "VARCHAR(16)"),
    ]
    missing = [(col, typ) for col, typ in additions if col not in existing]
    if not missing:
        return
    with engine.begin() as conn:
        for col, typ in missing:
            conn.execute(text(f"ALTER TABLE model_config ADD COLUMN {col} {typ}"))


def ensure_parse_record_columns():
    """为已存在的 parse_records 表补齐软失效相关列（create_all 不会给旧表加列）。
    新增列为 nullable，NULL 视为有效（未失效）。"""
    try:
        insp = inspect(engine)
        if not insp.has_table("parse_records"):
            return
        existing = {c["name"] for c in insp.get_columns("parse_records")}
    except Exception:
        return
    additions = [
        ("is_invalid", "BOOLEAN"),
        ("invalidated_at", "DATETIME"),
    ]
    missing = [(col, typ) for col, typ in additions if col not in existing]
    if not missing:
        return
    with engine.begin() as conn:
        for col, typ in missing:
            conn.execute(text(f"ALTER TABLE parse_records ADD COLUMN {col} {typ}"))
