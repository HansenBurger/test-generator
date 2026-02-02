"""
数据库模型定义
"""
from datetime import datetime
from sqlalchemy import Column, DateTime, Integer, String, Text

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
