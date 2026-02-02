"""
数据库访问封装
"""
from contextlib import contextmanager
from datetime import datetime
from typing import Optional, List

from app.db.database import SessionLocal
from app.db.models import ParseRecord, GenerationRecord


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
    with get_session() as session:
        return session.query(ParseRecord).filter(ParseRecord.outline_hash == outline_hash).first()


def get_parse_record_by_version_time(
    requirement_name: str,
    version: Optional[str],
    outline_time: Optional[str]
) -> Optional[ParseRecord]:
    with get_session() as session:
        return (
            session.query(ParseRecord)
            .filter(ParseRecord.requirement_name == requirement_name)
            .filter(ParseRecord.version == version)
            .filter(ParseRecord.outline_time == outline_time)
            .first()
        )


def get_parse_record(parse_id: str) -> Optional[ParseRecord]:
    with get_session() as session:
        return session.query(ParseRecord).filter(ParseRecord.parse_id == parse_id).first()


def list_parse_records(requirement_name: str) -> List[ParseRecord]:
    with get_session() as session:
        return (
            session.query(ParseRecord)
            .filter(ParseRecord.requirement_name == requirement_name)
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
