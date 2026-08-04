"""
数据库访问封装
"""
from contextlib import contextmanager
from datetime import datetime
from typing import Optional, List

from sqlalchemy import or_

from app.db.database import SessionLocal
from app.db.models import ParseRecord, GenerationRecord, ModelConfig


def _parse_record_valid_filter():
    """有效解析记录过滤条件：is_invalid 为 NULL 或 False 均视为有效"""
    return or_(ParseRecord.is_invalid.is_(None), ParseRecord.is_invalid == False)  # noqa: E712


@contextmanager
def get_session():
    session = SessionLocal()
    try:
        yield session
        session.commit()
    except Exception:
        session.rollback()
        raise
    finally:
        session.close()


def get_parse_record_by_hash(outline_hash: str) -> Optional[ParseRecord]:
    # 已失效的解析缓存不参与复用
    with get_session() as session:
        return (
            session.query(ParseRecord)
            .filter(ParseRecord.outline_hash == outline_hash)
            .filter(_parse_record_valid_filter())
            .first()
        )


def get_parse_record_by_version_time(
    requirement_name: str,
    version: Optional[str],
    outline_time: Optional[str]
) -> Optional[ParseRecord]:
    # 已失效记录不参与"同版本号+时间"的覆盖流程，重新上传会生成新记录
    with get_session() as session:
        return (
            session.query(ParseRecord)
            .filter(ParseRecord.requirement_name == requirement_name)
            .filter(ParseRecord.version == version)
            .filter(ParseRecord.outline_time == outline_time)
            .filter(_parse_record_valid_filter())
            .first()
        )


def get_parse_record(parse_id: str) -> Optional[ParseRecord]:
    with get_session() as session:
        return session.query(ParseRecord).filter(ParseRecord.parse_id == parse_id).first()


def list_parse_records(requirement_name: str) -> List[ParseRecord]:
    # 历史版本列表仅展示有效记录
    with get_session() as session:
        return (
            session.query(ParseRecord)
            .filter(ParseRecord.requirement_name == requirement_name)
            .filter(_parse_record_valid_filter())
            .order_by(ParseRecord.upload_time.desc())
            .all()
        )


def create_parse_record(
    parse_id: str,
    requirement_name: str,
    version: Optional[str],
    outline_time: Optional[str],
    outline_hash: str,
    status: str
) -> ParseRecord:
    with get_session() as session:
        record = ParseRecord(
            parse_id=parse_id,
            requirement_name=requirement_name,
            version=version,
            outline_time=outline_time,
            outline_hash=outline_hash,
            status=status,
            upload_time=datetime.utcnow()
        )
        session.add(record)
        session.flush()
        return record


def update_parse_record(
    parse_id: str,
    status: Optional[str] = None,
    test_point_count: Optional[int] = None,
    json_path: Optional[str] = None,
    xmind_path: Optional[str] = None
) -> Optional[ParseRecord]:
    with get_session() as session:
        record = session.query(ParseRecord).filter(ParseRecord.parse_id == parse_id).first()
        if not record:
            return None
        if status is not None:
            record.status = status
        if test_point_count is not None:
            record.test_point_count = test_point_count
        if json_path is not None:
            record.json_path = json_path
        if xmind_path is not None:
            record.xmind_path = xmind_path
        session.flush()
        return record


def update_parse_record_outline_hash(
    parse_id: str,
    outline_hash: str,
    upload_time: Optional[datetime] = None,
) -> Optional[ParseRecord]:
    """更新解析记录的哈希（可选更新上传时间）"""
    with get_session() as session:
        record = session.query(ParseRecord).filter(ParseRecord.parse_id == parse_id).first()
        if not record:
            return None
        record.outline_hash = outline_hash
        if upload_time is not None:
            record.upload_time = upload_time
        session.flush()
        return record


def invalidate_parse_record(parse_id: str) -> Optional[ParseRecord]:
    """将单条解析记录标记为失效（软删除）"""
    with get_session() as session:
        record = session.query(ParseRecord).filter(ParseRecord.parse_id == parse_id).first()
        if not record:
            return None
        record.is_invalid = True
        record.invalidated_at = datetime.utcnow()
        session.flush()
        return record


def invalidate_all_parse_records() -> int:
    """将所有未失效的解析记录标记为失效（软删除），返回本次打标数量。

    失效后：哈希/版本号缓存复用与历史版本列表均不再命中这些记录，
    重新上传同一文件会重新解析并生成新记录；旧记录与缓存 JSON 保留用于追溯。
    """
    with get_session() as session:
        count = (
            session.query(ParseRecord)
            .filter(_parse_record_valid_filter())
            .update(
                {ParseRecord.is_invalid: True, ParseRecord.invalidated_at: datetime.utcnow()},
                synchronize_session=False,
            )
        )
        return count


def get_generation_record(session_id: str) -> Optional[GenerationRecord]:
    with get_session() as session:
        return session.query(GenerationRecord).filter(GenerationRecord.session_id == session_id).first()


def create_or_update_generation_record(
    session_id: str,
    parse_record_id: str,
    prompt_strategy: Optional[str],
    prompt_version: Optional[str],
    generation_mode: Optional[str],
    status: str
) -> GenerationRecord:
    with get_session() as session:
        record = session.query(GenerationRecord).filter(GenerationRecord.session_id == session_id).first()
        if record:
            record.parse_record_id = parse_record_id
            record.prompt_strategy = prompt_strategy
            record.prompt_version = prompt_version
            record.generation_mode = generation_mode
            record.status = status
            record.start_time = datetime.utcnow()
            record.success_count = 0
            record.fail_count = 0
            record.json_path = None
            record.xmind_path = None
            record.completed_at = None
        else:
            record = GenerationRecord(
                session_id=session_id,
                parse_record_id=parse_record_id,
                prompt_strategy=prompt_strategy,
                prompt_version=prompt_version,
                generation_mode=generation_mode,
                status=status,
                start_time=datetime.utcnow()
            )
            session.add(record)
        session.flush()
        return record


def update_generation_record(
    session_id: str,
    status: Optional[str] = None,
    success_count: Optional[int] = None,
    fail_count: Optional[int] = None,
    json_path: Optional[str] = None,
    xmind_path: Optional[str] = None,
    completed_at: Optional[datetime] = None
) -> Optional[GenerationRecord]:
    with get_session() as session:
        record = session.query(GenerationRecord).filter(GenerationRecord.session_id == session_id).first()
        if not record:
            return None
        if status is not None:
            record.status = status
        if success_count is not None:
            record.success_count = success_count
        if fail_count is not None:
            record.fail_count = fail_count
        if json_path is not None:
            record.json_path = json_path
        if xmind_path is not None:
            record.xmind_path = xmind_path
        if completed_at is not None:
            record.completed_at = completed_at
        session.flush()
        return record


def get_model_config() -> Optional[ModelConfig]:
    """读取运行时模型配置（单行），不存在返回 None。"""
    with get_session() as session:
        return session.query(ModelConfig).filter(ModelConfig.id == 1).first()


def upsert_model_config(
    enable_thinking: bool,
    thinking_token_buffer: int,
    current_model: Optional[str],
    temperature: Optional[float] = None,
    model_mode: Optional[str] = None,
) -> ModelConfig:
    """全量写入运行时模型配置；current_model/temperature/model_mode 传 None 表示沿用默认。"""
    with get_session() as session:
        record = session.query(ModelConfig).filter(ModelConfig.id == 1).first()
        if not record:
            record = ModelConfig(id=1)
            session.add(record)
        record.enable_thinking = enable_thinking
        record.thinking_token_buffer = thinking_token_buffer
        record.current_model = current_model
        record.temperature = temperature
        record.model_mode = model_mode
        session.flush()
        return record
