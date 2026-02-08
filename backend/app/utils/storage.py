"""
文件存储工具
"""
import json
from typing import Any

from app.core.config import (
    DATA_DIR,
    PARSED_DIR,
    GENERATION_DIR,
    XMIND_DIR,
    ensure_dir
)


def get_base_data_dir() -> str:
    return ensure_dir(DATA_DIR)


def get_parsed_dir() -> str:
    return ensure_dir(PARSED_DIR)


def get_generation_dir() -> str:
    return ensure_dir(GENERATION_DIR)


def get_xmind_dir() -> str:
    return ensure_dir(XMIND_DIR)


def save_json(path: str, data: Any):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
