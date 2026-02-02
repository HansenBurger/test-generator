"""
文件存储工具
"""
import json
import os
from typing import Any


def _ensure_dir(path: str) -> str:
    os.makedirs(path, exist_ok=True)
    return path


def get_base_data_dir() -> str:
    base_dir = os.path.abspath(os.path.join(os.path.dirname(__file__), "..", ".."))
    return _ensure_dir(os.path.join(base_dir, "data"))


def get_parsed_dir() -> str:
    return _ensure_dir(os.path.join(get_base_data_dir(), "parsed"))


def get_generation_dir() -> str:
    return _ensure_dir(os.path.join(get_base_data_dir(), "generation"))


def get_xmind_dir() -> str:
    return _ensure_dir(os.path.join(get_base_data_dir(), "xmind"))


def save_json(path: str, data: Any):
    with open(path, "w", encoding="utf-8") as f:
        json.dump(data, f, ensure_ascii=False, indent=2)
