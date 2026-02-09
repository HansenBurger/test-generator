"""
配置项
"""
import os
from urllib.parse import quote_plus


def ensure_dir(path: str) -> str:
    os.makedirs(path, exist_ok=True)
    return path


BASE_DIR = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
DEFAULT_DATA_DIR = os.path.join(BASE_DIR, "data")

DATA_DIR = os.getenv("DATA_DIR", DEFAULT_DATA_DIR)
PARSED_DIR = os.getenv("PARSED_DIR", os.path.join(DATA_DIR, "parsed"))
GENERATION_DIR = os.getenv("GENERATION_DIR", os.path.join(DATA_DIR, "generation"))
XMIND_DIR = os.getenv("XMIND_DIR", os.path.join(DATA_DIR, "xmind"))

DASHSCOPE_API_BASE_URL = (
    os.getenv("DASHSCOPE_API_BASE_URL", "").strip()
    or "https://dashscope.aliyuncs.com/compatible-mode/v1"
)

DATABASE_URL = os.getenv("DATABASE_URL", "").strip()

if not DATABASE_URL:
    db_host = os.getenv("DB_HOST", "").strip()
    if db_host:
        db_port = os.getenv("DB_PORT", "3306").strip()
        db_user = os.getenv("DB_USER", "root").strip()
        db_password = os.getenv("DB_PASSWORD", "").strip()
        db_name = os.getenv("DB_NAME", "test_generator").strip()
        password = quote_plus(db_password) if db_password else ""
        auth = f"{db_user}:{password}" if password or db_user else db_user
        DATABASE_URL = f"mysql+pymysql://{auth}@{db_host}:{db_port}/{db_name}?charset=utf8mb4"
    else:
        ensure_dir(DATA_DIR)
        sqlite_path = os.path.join(DATA_DIR, "app.db")
        DATABASE_URL = f"sqlite:///{sqlite_path}"

SQLITE_CONNECT_ARGS = {"check_same_thread": False} if DATABASE_URL.startswith("sqlite") else {}
