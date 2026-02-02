"""
测试用例生成服务
"""
import json
import os
import threading
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional, Tuple
from uuid import uuid4

from app.models.schemas import ParsedXmindDocument, TestPoint, TestCase
from app.services.ai_client import AIClient
from app.db import repository
from app.utils.logger import generator_logger
from app.utils.storage import get_generation_dir, save_json


PROMPT_VERSION = "v1"
BATCH_SIZE = 80
ENABLE_BATCH = False

SYSTEM_PROMPT_METADATA = """你是一个银行信贷业务测试专家，需要分析测试点文本，补充缺失的正例反例标志和优先级标志。

输入：测试点文本
输出：JSON格式，包含：
- subtype: "positive"或"negative"
- priority: 1、2、3（1=高，2=中，3=低）

判断规则：
1. 正例判断关键词：
   - 正例："通过"、"成功"、"正确"、"一致"、"正常"
   - 反例："不通过"、"失败"、"错误"、"不一致"、"异常"、"提示"
2. 优先级判断：
   - 高优先级（1）：核心业务流程、关键检查规则、主要处理规则
   - 中优先级（2）：普通业务规则、重要但不是关键
   - 低优先级（3）：辅助性规则、边界条件
3. 如果文本中同时出现正反例关键词，以最终结果为准
"""

SYSTEM_PROMPT_METADATA_BATCH = """你是一个银行信贷业务测试专家，需要分析测试点文本，补充缺失的正例反例标志和优先级标志。

输入：JSON数组，每项包含：
- point_id: 测试点ID
- text: 测试点文本

输出：JSON数组（顺序与输入一致），每项包含：
- point_id
- subtype: "positive"或"negative"
- priority: 1、2、3（1=高，2=中，3=低）
"""

SYSTEM_PROMPT_CASE = """你是银行信贷项目测试专家，需要根据测试点生成可执行的测试用例。

输出 JSON，字段如下：
- preconditions: 前提条件数组
- steps: 测试步骤数组
- expected_results: 预期结果数组

要求：
1. 每个数组元素为一句话，简洁、可执行
2. 保持与测试点语义一致，不要引入无关内容
3. 不要输出多余解释或 Markdown
"""

SYSTEM_PROMPT_CASE_BATCH = """你是银行信贷项目测试专家，需要根据测试点生成可执行的测试用例。

输入：JSON数组，每项包含：
- point_id
- point_type
- subtype
- priority
- text
- flow_steps: 相关流程步骤（数组，可为空，仅用于规则类测试点）

输出：JSON数组（顺序与输入一致），每项包含：
- point_id
- preconditions: 前提条件数组
- steps: 测试步骤数组
- expected_results: 预期结果数组

要求：
1. 每个数组元素为一句话，简洁、可执行
2. 保持与测试点语义一致，不要引入无关内容
3. 不要输出多余解释或 Markdown
"""

def detect_subtype(text: str) -> Optional[str]:
    positive_keywords = ["通过", "成功", "正确", "一致", "正常"]
    negative_keywords = ["不通过", "失败", "错误", "不一致", "异常", "提示"]
    last_pos = -1
    subtype = None
    for word in positive_keywords:
        idx = text.rfind(word)
        if idx > last_pos:
            last_pos = idx
            subtype = "positive"
    for word in negative_keywords:
        idx = text.rfind(word)
        if idx > last_pos:
            last_pos = idx
            subtype = "negative"
    return subtype


def select_preview_points(points: List[TestPoint], count: Optional[int]) -> List[TestPoint]:
    if not points:
        return []
    target = count or 4
    target = max(3, min(5, target))

    def sort_key(p: TestPoint):
        priority_rank = {1: 0, 2: 1, 3: 2}.get(p.priority, 3)
        return (priority_rank, p.point_type, p.subtype or "")

    sorted_points = sorted(points, key=sort_key)

    selected: List[TestPoint] = []
    for point in sorted_points:
        if len(selected) >= target:
            break
        selected.append(point)

    return selected


def _normalize_list(value) -> List[str]:
    if value is None:
        return []
    if isinstance(value, list):
        return [str(v).strip() for v in value if str(v).strip()]
    if isinstance(value, str):
        return [v.strip() for v in value.split("\n") if v.strip()]
    return [str(value).strip()]


def _build_case_id() -> str:
    return f"TC{uuid4().hex[:8].upper()}"


def _build_user_prompt(point: TestPoint) -> str:
    return json.dumps({
        "point_id": point.point_id,
        "point_type": point.point_type,
        "subtype": point.subtype,
        "priority": point.priority,
        "text": point.text
    }, ensure_ascii=False)

def _normalize_context_key(context: Optional[str]) -> str:
    value = context or ""
    return value.replace("业务规则", "业务流程").strip()


def _chunk_list(items: List[TestPoint], size: int) -> List[List[TestPoint]]:
    if size <= 0:
        return [items]
    return [items[i:i + size] for i in range(0, len(items), size)]


class CaseGenerator:
    """测试用例生成器"""

    def __init__(self):
        self._client = AIClient()

    def fill_missing_metadata(self, points: List[TestPoint]) -> Tuple[int, List[str]]:
        token_usage = 0
        logs: List[str] = []
        missing = [p for p in points if not (p.subtype and p.priority)]
        if not missing:
            return token_usage, logs
        for batch in _chunk_list(missing, BATCH_SIZE):
            payload = [{"point_id": p.point_id, "text": p.text} for p in batch]
            try:
                result, tokens = self._client.chat_json(
                    system_prompt=SYSTEM_PROMPT_METADATA_BATCH,
                    user_prompt=json.dumps(payload, ensure_ascii=False),
                    temperature=0.1,
                    max_tokens=800
                )
                token_usage += tokens
                if isinstance(result, list):
                    result_map = {str(item.get("point_id")): item for item in result if isinstance(item, dict)}
                    for point in batch:
                        meta = result_map.get(point.point_id, {})
                        point.subtype = point.subtype or meta.get("subtype")
                        point.priority = point.priority or meta.get("priority")
                else:
                    logs.append("元数据批量补全失败：返回格式非数组")
            except Exception as exc:
                logs.append(f"元数据批量补全失败：{str(exc)}")
        return token_usage, logs

    def generate_case(self, point: TestPoint, strategy: str = "standard") -> Tuple[TestCase, int]:
        temperature = 0.2 if strategy == "standard" else 0.6
        max_tokens = 900 if strategy == "standard" else 600
        result, tokens = self._client.chat_json(
            system_prompt=SYSTEM_PROMPT_CASE,
            user_prompt=_build_user_prompt(point),
            temperature=temperature,
            max_tokens=max_tokens
        )
        case = TestCase(
            case_id=_build_case_id(),
            point_id=point.point_id,
            point_type=point.point_type,
            subtype=point.subtype,
            priority=point.priority,
            text=point.text,
            preconditions=_normalize_list(result.get("preconditions")),
            steps=_normalize_list(result.get("steps")),
            expected_results=_normalize_list(result.get("expected_results"))
        )
        return case, tokens

    def generate_cases_batch(
        self,
        points: List[TestPoint],
        strategy: str,
        flow_steps_map: Optional[Dict[str, List[str]]] = None
    ) -> Tuple[List[TestCase], int, List[str]]:
        if not ENABLE_BATCH:
            cases: List[TestCase] = []
            logs: List[str] = []
            token_usage = 0
            for point in points:
                try:
                    case, tokens = self.generate_case(point, strategy)
                    cases.append(case)
                    token_usage += tokens
                except Exception as exc:
                    logs.append(f"测试点 {point.point_id} 生成失败：{str(exc)}")
            return cases, token_usage, logs
        temperature = 0.2 if strategy == "standard" else 0.6
        max_tokens = 1200 if strategy == "standard" else 900
        payload = []
        for point in points:
            flow_steps = []
            if point.point_type == "rule" and flow_steps_map is not None:
                key = _normalize_context_key(point.context)
                flow_steps = flow_steps_map.get(key, [])
            payload.append({
                "point_id": point.point_id,
                "point_type": point.point_type,
                "subtype": point.subtype,
                "priority": point.priority,
                "text": point.text,
                "flow_steps": flow_steps
            })
        logs: List[str] = []
        try:
            result, tokens = self._client.chat_json(
                system_prompt=SYSTEM_PROMPT_CASE_BATCH,
                user_prompt=json.dumps(payload, ensure_ascii=False),
                temperature=temperature,
                max_tokens=max_tokens
            )
        except Exception as exc:
            logs: List[str] = [f"批量生成解析失败，降级单点生成：{str(exc)}"]
            cases: List[TestCase] = []
            token_usage = 0
            for point in points:
                single_payload = [{
                    "point_id": point.point_id,
                    "point_type": point.point_type,
                    "subtype": point.subtype,
                    "priority": point.priority,
                    "text": point.text,
                    "flow_steps": flow_steps_map.get(_normalize_context_key(point.context), [])
                    if point.point_type == "rule" and flow_steps_map else []
                }]
                try:
                    single_result, single_tokens = self._client.chat_json(
                        system_prompt=SYSTEM_PROMPT_CASE_BATCH,
                        user_prompt=json.dumps(single_payload, ensure_ascii=False),
                        temperature=temperature,
                        max_tokens=max_tokens
                    )
                    token_usage += single_tokens
                    if isinstance(single_result, list) and single_result:
                        item = single_result[0]
                        case = TestCase(
                            case_id=_build_case_id(),
                            point_id=point.point_id,
                            point_type=point.point_type,
                            subtype=point.subtype,
                            priority=point.priority,
                            text=point.text,
                            preconditions=_normalize_list(item.get("preconditions")),
                            steps=_normalize_list(item.get("steps")),
                            expected_results=_normalize_list(item.get("expected_results"))
                        )
                        cases.append(case)
                    else:
                        logs.append(f"测试点 {point.point_id} 生成结果缺失")
                except Exception as inner_exc:
                    logs.append(f"测试点 {point.point_id} 生成失败：{str(inner_exc)}")
            return cases, token_usage, logs
        cases: List[TestCase] = []
        if not isinstance(result, list):
            logs.append("批量生成返回格式异常，非数组")
            return cases, tokens, logs
        result_map = {str(item.get("point_id")): item for item in result if isinstance(item, dict)}
        for point in points:
            item = result_map.get(point.point_id)
            if not item:
                logs.append(f"测试点 {point.point_id} 生成结果缺失")
                continue
            case = TestCase(
                case_id=_build_case_id(),
                point_id=point.point_id,
                point_type=point.point_type,
                subtype=point.subtype,
                priority=point.priority,
                text=point.text,
                preconditions=_normalize_list(item.get("preconditions")),
                steps=_normalize_list(item.get("steps")),
                expected_results=_normalize_list(item.get("expected_results"))
            )
            cases.append(case)
        return cases, tokens, logs


@dataclass
class GenerationTask:
    task_id: str
    requirement_name: str
    strategy: str
    parse_id: Optional[str] = None
    session_id: Optional[str] = None
    prompt_version: Optional[str] = None
    generation_mode: Optional[str] = None
    points: List[TestPoint] = field(default_factory=list)
    status: str = "pending"
    progress: float = 0.0
    total: int = 0
    completed: int = 0
    failed: int = 0
    logs: List[str] = field(default_factory=list)
    cases: List[TestCase] = field(default_factory=list)
    token_usage: int = 0
    error: Optional[str] = None
    created_at: datetime = field(default_factory=datetime.now)
    started_at: Optional[datetime] = None
    completed_at: Optional[datetime] = None


class CaseGenerationManager:
    """生成任务管理器"""

    def __init__(self):
        self._parsed_docs: Dict[str, ParsedXmindDocument] = {}
        self._parsed_meta: Dict[str, datetime] = {}
        self._preview_cache: Dict[str, Dict] = {}
        self._tasks: Dict[str, GenerationTask] = {}
        self._lock = threading.Lock()
        self._executor = ThreadPoolExecutor(max_workers=3, thread_name_prefix="case_generator")
        self._generator = CaseGenerator()

    def save_parsed_doc(self, parsed: ParsedXmindDocument):
        with self._lock:
            self._parsed_docs[parsed.parse_id] = parsed
            self._parsed_meta[parsed.parse_id] = datetime.now()

    def get_parsed_doc(self, parse_id: str) -> Optional[ParsedXmindDocument]:
        with self._lock:
            cached = self._parsed_docs.get(parse_id)
            if cached:
                return cached
        record = repository.get_parse_record(parse_id)
        if not record or not record.json_path:
            return None
        try:
            with open(record.json_path, "r", encoding="utf-8") as f:
                data = json.load(f)
            parsed = ParsedXmindDocument(**data)
            with self._lock:
                self._parsed_docs[parse_id] = parsed
                self._parsed_meta[parse_id] = record.upload_time
            return parsed
        except Exception:
            return None

    def list_versions(self, requirement_name: str) -> List[Dict[str, str]]:
        result: List[Dict[str, str]] = []
        records = repository.list_parse_records(requirement_name)
        for record in records:
            result.append({
                "parse_id": record.parse_id,
                "created_at": record.upload_time.isoformat() if record.upload_time else ""
            })
        return result

    def create_preview(self, parse_id: str, cases: List[TestCase], point_ids: List[str]) -> str:
        preview_id = uuid4().hex
        with self._lock:
            self._preview_cache[preview_id] = {
                "parse_id": parse_id,
                "cases": cases,
                "point_ids": point_ids
            }
        return preview_id

    def get_preview(self, preview_id: str) -> Optional[Dict]:
        with self._lock:
            return self._preview_cache.get(preview_id)

    def generate_preview(self, parse_id: str, count: Optional[int]) -> Tuple[str, List[TestCase], int, List[str]]:
        parsed = self.get_parsed_doc(parse_id)
        if not parsed:
            raise ValueError("解析结果不存在，请先上传并解析XMind")

        token_usage, logs = self._generator.fill_missing_metadata(parsed.test_points)
        selected_points = select_preview_points(parsed.test_points, count)

        cases: List[TestCase] = []
        for point in selected_points:
            case, tokens = self._generator.generate_case(point, "standard")
            cases.append(case)
            token_usage += tokens

        preview_id = self.create_preview(parse_id, cases, [p.point_id for p in selected_points])
        return preview_id, cases, token_usage, logs

    def create_task_from_preview(
        self,
        preview_id: str,
        strategy: str,
        session_id: Optional[str] = None,
        prompt_version: Optional[str] = None
    ) -> Tuple[str, List[TestCase], str]:
        preview = self.get_preview(preview_id)
        if not preview:
            raise ValueError("预生成记录不存在")

        parse_id = preview["parse_id"]
        parsed = self.get_parsed_doc(parse_id)
        if not parsed:
            raise ValueError("解析结果不存在，请重新上传")

        preview_point_ids = set(preview.get("point_ids", []))
        remaining_points = [p for p in parsed.test_points if p.point_id not in preview_point_ids]
        task_id, session_id = self.create_generation_task(
            requirement_name=parsed.requirement_name,
            parse_id=parse_id,
            points=remaining_points,
            strategy=strategy,
            initial_cases=preview.get("cases", []),
            session_id=session_id,
            prompt_version=prompt_version,
            generation_mode="preview"
        )
        return task_id, preview.get("cases", []), session_id

    def create_task_for_parse(
        self,
        parse_id: str,
        strategy: str,
        session_id: Optional[str] = None,
        prompt_version: Optional[str] = None
    ) -> Tuple[str, str]:
        parsed = self.get_parsed_doc(parse_id)
        if not parsed:
            raise ValueError("解析结果不存在，请先上传并解析XMind")
        return self.create_generation_task(
            requirement_name=parsed.requirement_name,
            parse_id=parse_id,
            points=parsed.test_points,
            strategy=strategy,
            initial_cases=[],
            session_id=session_id,
            prompt_version=prompt_version,
            generation_mode="bulk"
        )

    def create_generation_task(
        self,
        requirement_name: str,
        parse_id: Optional[str],
        points: List[TestPoint],
        strategy: str,
        initial_cases: Optional[List[TestCase]] = None,
        session_id: Optional[str] = None,
        prompt_version: Optional[str] = None,
        generation_mode: Optional[str] = None
    ) -> Tuple[str, str]:
        task_id = uuid4().hex
        session_id = session_id or uuid4().hex
        task = GenerationTask(
            task_id=task_id,
            requirement_name=requirement_name,
            parse_id=parse_id,
            session_id=session_id,
            strategy=strategy,
            prompt_version=prompt_version or PROMPT_VERSION,
            generation_mode=generation_mode,
            points=points,
            total=len(points)
        )
        if initial_cases:
            task.cases.extend(initial_cases)

        with self._lock:
            self._tasks[task_id] = task

        if parse_id:
            repository.create_or_update_generation_record(
                session_id=session_id,
                parse_record_id=parse_id,
                prompt_strategy=strategy,
                prompt_version=task.prompt_version,
                generation_mode=generation_mode,
                status="pending"
            )

        self._executor.submit(self._run_task, task_id)
        return task_id, session_id

    def get_task(self, task_id: str) -> Optional[GenerationTask]:
        with self._lock:
            return self._tasks.get(task_id)

    def _run_task(self, task_id: str):
        task = self.get_task(task_id)
        if not task:
            return
        task.status = "processing"
        task.started_at = datetime.now()
        generator_logger.info(f"测试用例生成任务开始 - 任务ID: {task_id}")
        if task.session_id:
            repository.update_generation_record(task.session_id, status="processing")

        try:
            token_usage, logs = self._generator.fill_missing_metadata(task.points)
            task.token_usage += token_usage
            task.logs.extend(logs)

            point_map = {p.point_id: p for p in task.points}
            process_points = [p for p in task.points if p.point_type == "process"]
            rule_points = [p for p in task.points if p.point_type == "rule"]

            processed_count = 0
            flow_steps_map: Dict[str, List[str]] = {}

            for batch in _chunk_list(process_points, BATCH_SIZE):
                cases, tokens, batch_logs = self._generator.generate_cases_batch(batch, task.strategy)
                task.token_usage += tokens
                task.logs.extend(batch_logs)
                for case in cases:
                    task.cases.append(case)
                    task.completed += 1
                    task.logs.append(f"测试点 {case.point_id} 生成完成")
                    point = point_map.get(case.point_id)
                    if point:
                        key = _normalize_context_key(point.context)
                        flow_steps_map.setdefault(key, []).extend(case.steps or [])
                failed_in_batch = len(batch) - len(cases)
                if failed_in_batch > 0:
                    task.failed += failed_in_batch
                processed_count += len(batch)
                task.progress = processed_count / max(1, task.total)

            for batch in _chunk_list(rule_points, BATCH_SIZE):
                cases, tokens, batch_logs = self._generator.generate_cases_batch(
                    batch,
                    task.strategy,
                    flow_steps_map=flow_steps_map
                )
                task.token_usage += tokens
                task.logs.extend(batch_logs)
                for case in cases:
                    task.cases.append(case)
                    task.completed += 1
                    task.logs.append(f"测试点 {case.point_id} 生成完成")
                failed_in_batch = len(batch) - len(cases)
                if failed_in_batch > 0:
                    task.failed += failed_in_batch
                processed_count += len(batch)
                task.progress = processed_count / max(1, task.total)

            task.status = "completed"
            task.completed_at = datetime.now()
            generator_logger.info(f"测试用例生成任务完成 - 任务ID: {task_id}")
            if task.session_id:
                json_path = self._save_cases_json(task.session_id, task.cases)
                repository.update_generation_record(
                    task.session_id,
                    status="completed",
                    success_count=task.completed,
                    fail_count=task.failed,
                    json_path=json_path,
                    completed_at=task.completed_at
                )
        except Exception as exc:
            task.status = "failed"
            task.error = str(exc)
            task.completed_at = datetime.now()
            generator_logger.error(f"测试用例生成任务失败 - 任务ID: {task_id}, 错误: {str(exc)}", exc_info=True)
            if task.session_id:
                repository.update_generation_record(
                    task.session_id,
                    status="failed",
                    success_count=task.completed,
                    fail_count=task.failed,
                    completed_at=task.completed_at
                )

    def _save_cases_json(self, session_id: str, cases: List[TestCase]) -> str:
        generation_dir = get_generation_dir()
        path = os.path.join(generation_dir, f"cases_{session_id}.json")
        save_json(path, [case.model_dump() for case in cases])
        return path


case_generation_manager = CaseGenerationManager()
