"""
数据库模型定义
"""
from datetime import datetime
from sqlalchemy import Boolean, Column, DateTime, Float, Integer, String, Text

from app.db.database import Base


class ParseRecord(Base):
    __tablename__ = "parse_records"

    parse_id = Column(String(64), primary_key=True, index=True)
    requirement_name = Column(String(255), nullable=False)
    version = Column(String(64), nullable=True)
    outline_time = Column(String(64), nullable=True)
    upload_time = Column(DateTime, default=datetime.utcnow, nullable=False)
    outline_hash = Column(String(64), index=True, nullable=False)
    status = Column(String(32), nullable=False, default="pending")
    test_point_count = Column(Integer, default=0, nullable=False)
    json_path = Column(Text, nullable=True)
    xmind_path = Column(Text, nullable=True)


class GenerationRecord(Base):
    __tablename__ = "generation_records"

    session_id = Column(String(64), primary_key=True, index=True)
    parse_record_id = Column(String(64), nullable=False, index=True)
    prompt_strategy = Column(String(64), nullable=True)
    prompt_version = Column(String(64), nullable=True)
    generation_mode = Column(String(64), nullable=True)
    user_feedback = Column(Text, nullable=True)
    user_id = Column(String(64), nullable=True)
    start_time = Column(DateTime, default=datetime.utcnow, nullable=False)
    status = Column(String(32), nullable=False, default="pending")
    success_count = Column(Integer, default=0, nullable=False)
    fail_count = Column(Integer, default=0, nullable=False)
    json_path = Column(Text, nullable=True)
    xmind_path = Column(Text, nullable=True)
    completed_at = Column(DateTime, nullable=True)


class ModelConfig(Base):
    """运行时模型配置（单行，id 固定为 1）。NULL 字段表示沿用 .env/默认值。"""
    __tablename__ = "model_config"

    id = Column(Integer, primary_key=True)
    enable_thinking = Column(Boolean, nullable=True)
    thinking_token_buffer = Column(Integer, nullable=True)
    current_model = Column(String(128), nullable=True)
    temperature = Column(Float, nullable=True)
    model_mode = Column(String(16), nullable=True)
    updated_at = Column(DateTime, default=datetime.utcnow, onupdate=datetime.utcnow)
