"""
SQLAlchemy 数据库初始化
"""
from sqlalchemy import create_engine, inspect, text
from sqlalchemy.orm import declarative_base, sessionmaker

from app.core.config import DATABASE_URL, SQLITE_CONNECT_ARGS
from app.utils.logger import logger

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


def _ensure_table_columns(table: str, additions: list):
    """为已存在的表补齐新增列（create_all 不会给旧表加列）。

    SQLite / MySQL 均支持 ALTER TABLE ADD COLUMN，新增列均为 nullable。
    若数据库账号无 ALTER 权限（如内网受限 MySQL 账号），不阻塞启动，
    仅打印需要手动执行的 DDL；但在补列完成前，涉及新列的功能会报错。
    """
    try:
        insp = inspect(engine)
        if not insp.has_table(table):
            return
        existing = {c["name"] for c in insp.get_columns(table)}
    except Exception:
        return
    missing = [(col, typ) for col, typ in additions if col not in existing]
    if not missing:
        return
    ddl_statements = [f"ALTER TABLE {table} ADD COLUMN {col} {typ};" for col, typ in missing]
    try:
        with engine.begin() as conn:
            for ddl in ddl_statements:
                conn.execute(text(ddl))
        logger.info(f"已为 {table} 表自动补齐新列: {', '.join(col for col, _ in missing)}")
    except Exception as exc:
        logger.warning(
            f"为 {table} 表自动补列失败（可能数据库账号无 ALTER 权限）: {exc}。"
            f"请让 DBA 或有权限的账号手动执行以下 SQL 后再使用相关功能: " + " ".join(ddl_statements)
        )


def ensure_model_config_columns():
    """为已存在的 model_config 表补齐新增列"""
    _ensure_table_columns("model_config", [
        ("temperature", "FLOAT"),
        ("model_mode", "VARCHAR(16)"),
    ])


def ensure_parse_record_columns():
    """为已存在的 parse_records 表补齐软失效相关列。
    新增列为 nullable，NULL 视为有效（未失效）。"""
    _ensure_table_columns("parse_records", [
        ("is_invalid", "BOOLEAN"),
        ("invalidated_at", "DATETIME"),
    ])
