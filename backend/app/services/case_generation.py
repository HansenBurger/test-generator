"""
测试用例生成服务
"""
import json
import os
import threading
import os as _os
from concurrent.futures import as_completed
from concurrent.futures import ThreadPoolExecutor
from dataclasses import dataclass, field
from datetime import datetime
from typing import Dict, List, Optional, Tuple
from uuid import uuid4

from app.models.schemas import ParsedXmindDocument, TestPoint, TestCase
from app.services.ai_client import AIClient
from app.services.prompts import (
    PROMPT_VERSION,
    SYSTEM_PROMPT_METADATA,
    SYSTEM_PROMPT_METADATA_BATCH,
    get_case_prompt,
    get_case_batch_prompt
)
from app.db import repository
from app.utils.logger import generator_logger
from app.utils.storage import get_generation_dir, save_json


BATCH_SIZE = 80
# 批量生成（一次请求多条用例）默认仅用于“规则”类测试点
ENABLE_PROCESS_BATCH = False
ENABLE_RULE_BATCH = (_os.getenv("ENABLE_RULE_BATCH", "1") == "1")
RULE_BATCH_SIZE = 20
RULE_BATCH_CONCURRENCY = int(_os.getenv("RULE_BATCH_CONCURRENCY", "3") or 3)
METADATA_BATCH_SIZE = 10
METADATA_BATCH_CONCURRENCY = int(_os.getenv("METADATA_BATCH_CONCURRENCY", "3") or 3)


def detect_subtype(text: str) -> Optional[str]:
    positive_keywords = ["通过", "成功", "正确", "一致", "正常"]
    negative_keywords = ["不通过", "失败", "错误", "不一致", "异常", "提示"]

    has_positive = any(word in text for word in positive_keywords)
    has_negative = any(word in text for word in negative_keywords)

    if has_positive and has_negative:
        return None
    if has_negative:
        return "negative"
    if has_positive:
        return "positive"
    return None


def _has_subtype_conflict(text: Optional[str]) -> bool:
    if not text:
        return False
    positive_keywords = ["通过", "成功", "正确", "一致", "正常"]
    negative_keywords = ["不通过", "失败", "错误", "不一致", "异常", "提示"]
    has_positive = any(word in text for word in positive_keywords)
    has_negative = any(word in text for word in negative_keywords)
    return has_positive and has_negative


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


def _build_case_example(case: TestCase, for_batch: bool, include_flow_steps: bool) -> str:
    input_payload = {
        "point_id": case.point_id,
        "point_type": case.point_type,
        "subtype": case.subtype,
        "priority": case.priority,
        "text": case.text
    }
    if include_flow_steps:
        input_payload["flow_steps"] = case.steps or []
    output_payload = {
        "preconditions": case.preconditions or [],
        "steps": case.steps or [],
        "expected_results": case.expected_results or []
    }
    if for_batch:
        input_text = json.dumps([input_payload], ensure_ascii=False, indent=2)
        output_text = json.dumps([{"point_id": case.point_id, **output_payload}], ensure_ascii=False, indent=2)
    else:
        input_text = json.dumps(input_payload, ensure_ascii=False, indent=2)
        output_text = json.dumps(output_payload, ensure_ascii=False, indent=2)
    return f"示例：\n输入：{input_text}\n输出：{output_text}"


def _build_prompt_examples(cases: List[TestCase]) -> Dict[str, str]:
    examples: Dict[str, str] = {}
    process_case = next((case for case in cases if case.point_type == "process"), None)
    rule_case = next((case for case in cases if case.point_type == "rule"), None)
    if process_case:
        examples["process"] = _build_case_example(process_case, for_batch=False, include_flow_steps=False)
        examples["process_batch"] = _build_case_example(process_case, for_batch=True, include_flow_steps=False)
    if rule_case:
        examples["rule"] = _build_case_example(rule_case, for_batch=False, include_flow_steps=False)
        examples["rule_batch"] = _build_case_example(rule_case, for_batch=True, include_flow_steps=True)
    return examples

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

    def _resolve_case_prompt(
        self,
        point_type: str,
        strategy: str,
        prompt_examples: Optional[Dict[str, str]] = None,
        for_batch: bool = False
    ) -> str:
        normalized = "process" if point_type == "process" else "rule"
        example_key = f"{normalized}_batch" if for_batch else normalized
        example = prompt_examples.get(example_key, "") if prompt_examples else ""
        if for_batch:
            return get_case_batch_prompt(normalized, strategy, example)
        return get_case_prompt(normalized, strategy, example)

    def _request_metadata_batch(
        self,
        batch: List[TestPoint],
        attempt: int = 1
    ) -> Tuple[Dict[str, dict], int, List[str]]:
        logs: List[str] = []
        payload = [{"point_id": p.point_id, "text": p.text} for p in batch]
        try:
            result, tokens = self._client.chat_json(
                system_prompt=SYSTEM_PROMPT_METADATA_BATCH,
                user_prompt=json.dumps(payload, ensure_ascii=False),
                temperature=0.1,
                max_tokens=800
            )
            if not isinstance(result, list):
                raise ValueError("元数据批量补全失败：返回格式非数组")
            result_map = {
                str(item.get("point_id")): item
                for item in result
                if isinstance(item, dict) and item.get("point_id") is not None
            }
            missing_points = [p for p in batch if p.point_id not in result_map]
            if missing_points and attempt < 2:
                logs.append(f"元数据批量补全缺失 {len(missing_points)} 条，尝试重试")
                retry_map, retry_tokens, retry_logs = self._request_metadata_batch(
                    missing_points,
                    attempt=attempt + 1
                )
                result_map.update(retry_map)
                tokens += retry_tokens
                logs.extend(retry_logs)
            elif missing_points:
                logs.append(f"元数据批量补全缺失 {len(missing_points)} 条，已达重试上限")
            return result_map, tokens, logs
        except Exception as exc:
            if len(batch) > 1:
                logs.append(f"元数据批量补全失败，拆分重试：{str(exc)}")
                mid = len(batch) // 2
                left_map, left_tokens, left_logs = self._request_metadata_batch(batch[:mid], attempt=attempt + 1)
                right_map, right_tokens, right_logs = self._request_metadata_batch(batch[mid:], attempt=attempt + 1)
                logs.extend(left_logs)
                logs.extend(right_logs)
                merged = {}
                merged.update(left_map)
                merged.update(right_map)
                return merged, left_tokens + right_tokens, logs
            if attempt < 2:
                logs.append(f"元数据单点补全失败，重试一次：{str(exc)}")
                return self._request_metadata_batch(batch, attempt=attempt + 1)
            logs.append(f"元数据单点补全失败：{str(exc)}")
            return {}, 0, logs

    def fill_missing_metadata(self, points: List[TestPoint]) -> Tuple[int, List[str]]:
        token_usage = 0
        logs: List[str] = []
        missing = [p for p in points if not (p.subtype and p.priority)]
        if not missing:
            return token_usage, logs
        if len(missing) > METADATA_BATCH_SIZE:
            batches = _chunk_list(missing, METADATA_BATCH_SIZE)
            concurrency = max(1, min(METADATA_BATCH_CONCURRENCY, len(batches)))
            logs.append(f"元数据补全：分批 {len(batches)} 组，每批 {METADATA_BATCH_SIZE}，并发 {concurrency}")
            with ThreadPoolExecutor(max_workers=concurrency, thread_name_prefix="metadata_batch") as pool:
                futures = {
                    pool.submit(self._request_metadata_batch, batch): batch
                    for batch in batches
                }
                for future in as_completed(futures):
                    batch = futures[future]
                    result_map, tokens, batch_logs = future.result()
                    token_usage += tokens
                    logs.extend(batch_logs)
                    for point in batch:
                        meta = result_map.get(point.point_id)
                        if not meta:
                            continue
                        point.subtype = point.subtype or meta.get("subtype")
                        point.priority = point.priority or meta.get("priority")
        else:
            for batch in _chunk_list(missing, METADATA_BATCH_SIZE):
                result_map, tokens, batch_logs = self._request_metadata_batch(batch)
                token_usage += tokens
                logs.extend(batch_logs)
                for point in batch:
                    meta = result_map.get(point.point_id, {})
                    point.subtype = point.subtype or meta.get("subtype")
                    point.priority = point.priority or meta.get("priority")
        return token_usage, logs

    def generate_case(
        self,
        point: TestPoint,
        strategy: str = "standard",
        prompt_examples: Optional[Dict[str, str]] = None
    ) -> Tuple[TestCase, int]:
        temperature = 0.2 if strategy == "standard" else 0.6
        max_tokens = 900 if strategy == "standard" else 600
        result, tokens = self._client.chat_json(
            system_prompt=self._resolve_case_prompt(point.point_type, strategy, prompt_examples, for_batch=False),
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
        flow_steps_map: Optional[Dict[str, List[str]]] = None,
        enable_batch: bool = False,
        prompt_examples: Optional[Dict[str, str]] = None
    ) -> Tuple[List[TestCase], int, List[str]]:
        if not enable_batch:
            cases: List[TestCase] = []
            logs: List[str] = []
            token_usage = 0
            failed_ids: List[str] = []
            for point in points:
                try:
                    case, tokens = self.generate_case(point, strategy, prompt_examples=prompt_examples)
                    cases.append(case)
                    token_usage += tokens
                except Exception as exc:
                    failed_ids.append(point.point_id)
            if failed_ids:
                sample = ", ".join(failed_ids[:5])
                more = "..." if len(failed_ids) > 5 else ""
                logs.append(f"批次内生成失败 {len(failed_ids)} 条（示例：{sample}{more}）")
            return cases, token_usage, logs
        # 批量模式更容易出现“输出格式漂移”，这里降温以提升 JSON 稳定性
        temperature = 0.1
        # 批量输出 token 需求与条数近似线性相关，这里做一个保守估计并设置上限
        base = 900 if strategy == "standard" else 700
        per_item = 70 if strategy == "standard" else 55
        max_tokens = min(6000, base + len(points) * per_item)
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
                system_prompt=self._resolve_case_prompt(
                    points[0].point_type if points else "process",
                    strategy,
                    prompt_examples,
                    for_batch=True
                ),
                user_prompt=json.dumps(payload, ensure_ascii=False),
                temperature=temperature,
                max_tokens=max_tokens
            )
        except Exception as exc:
            # 批量失败的根因通常是模型输出 JSON 不合法（缺逗号/括号/截断等）
            # 先“拆批重试”（更符合你的预期：分段异步），实在不行再降级单点
            logs = [f"批量生成解析失败，拆分重试：{str(exc)}"]
            if len(points) > 1:
                mid = len(points) // 2
                left = points[:mid]
                right = points[mid:]
                left_cases, left_tokens, left_logs = self.generate_cases_batch(
                    left,
                    strategy,
                    flow_steps_map=flow_steps_map,
                    enable_batch=True,
                    prompt_examples=prompt_examples
                )
                right_cases, right_tokens, right_logs = self.generate_cases_batch(
                    right,
                    strategy,
                    flow_steps_map=flow_steps_map,
                    enable_batch=True,
                    prompt_examples=prompt_examples
                )
                return left_cases + right_cases, left_tokens + right_tokens, logs + left_logs + right_logs

            # 只有 1 条仍失败：改用单点 prompt 生成（更稳）
            try:
                case, single_tokens = self.generate_case(points[0], strategy, prompt_examples=prompt_examples)
                return [case], single_tokens, logs + ["单点降级成功"]
            except Exception as inner_exc:
                return [], 0, logs + [f"单点降级仍失败：{str(inner_exc)}"]
        cases: List[TestCase] = []
        if not isinstance(result, list):
            logs.append("批量生成返回格式异常，非数组")
            return cases, tokens, logs
        result_map = {str(item.get("point_id")): item for item in result if isinstance(item, dict)}
        missing_ids: List[str] = []
        for point in points:
            item = result_map.get(point.point_id)
            if not item:
                missing_ids.append(point.point_id)
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
        if missing_ids:
            sample = ", ".join(missing_ids[:5])
            more = "..." if len(missing_ids) > 5 else ""
            logs.append(f"批量生成结果缺失 {len(missing_ids)} 条（示例：{sample}{more}）")
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
    prompt_examples: Dict[str, str] = field(default_factory=dict)
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

        # 预生成属于“生成”范畴：跳过 priority=3 的测试点
        generatable_points = [p for p in parsed.test_points if p.priority != 3]
        if not generatable_points:
            raise ValueError("无可生成的测试点（均为低优先级 priority=3）")

        token_usage, logs = self._generator.fill_missing_metadata(generatable_points)
        selected_points = select_preview_points(generatable_points, count)

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
        remaining_points = [
            p for p in parsed.test_points
            if p.point_id not in preview_point_ids and p.priority != 3
        ]
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
        points = [p for p in parsed.test_points if p.priority != 3]
        return self.create_generation_task(
            requirement_name=parsed.requirement_name,
            parse_id=parse_id,
            points=points,
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
        prompt_examples: Dict[str, str] = {}
        if generation_mode == "preview" and initial_cases:
            prompt_examples = _build_prompt_examples(initial_cases)
        task = GenerationTask(
            task_id=task_id,
            requirement_name=requirement_name,
            parse_id=parse_id,
            session_id=session_id,
            strategy=strategy,
            prompt_version=prompt_version or PROMPT_VERSION,
            generation_mode=generation_mode,
            prompt_examples=prompt_examples,
            points=points,
            total=len(points)
        )
        if initial_cases:
            task.cases.extend(initial_cases)
            task.completed = len(initial_cases)
            task.total = len(points) + len(initial_cases)
            task.progress = (task.completed / max(1, task.total))

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

    def get_task_by_session(self, session_id: str) -> Optional[GenerationTask]:
        if not session_id:
            return None
        with self._lock:
            for task in self._tasks.values():
                if task.session_id == session_id:
                    return task
        return None

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
            # 预处理：补全元数据
            conflict_count = sum(1 for p in task.points if _has_subtype_conflict(p.text))
            missing_count = len([p for p in task.points if not (p.subtype and p.priority)])
            task.logs.append(f"阶段：预处理（元数据补全），待补全 {missing_count} 条")
            task.logs.append(f"正反例关键词冲突：{conflict_count} 条")
            token_usage, logs = self._generator.fill_missing_metadata(task.points)
            task.token_usage += token_usage
            task.logs.extend(logs)
            task.logs.append("阶段：预处理完成")

            point_map = {p.point_id: p for p in task.points}
            process_points = [p for p in task.points if p.point_type == "process"]
            rule_points = [p for p in task.points if p.point_type == "rule"]

            flow_steps_map: Dict[str, List[str]] = {}

            if process_points:
                process_batches = _chunk_list(process_points, BATCH_SIZE)
                task.logs.append(f"阶段：流程用例生成，共 {len(process_points)} 条，批次数 {len(process_batches)}")
            else:
                process_batches = []
                task.logs.append("阶段：流程用例生成，0 条（跳过）")

            for idx, batch in enumerate(process_batches, start=1):
                cases, tokens, batch_logs = self._generator.generate_cases_batch(
                    batch,
                    task.strategy,
                    enable_batch=ENABLE_PROCESS_BATCH,
                    prompt_examples=task.prompt_examples
                )
                task.token_usage += tokens
                task.logs.extend(batch_logs)
                for case in cases:
                    task.cases.append(case)
                    task.completed += 1
                    point = point_map.get(case.point_id)
                    if point:
                        key = _normalize_context_key(point.context)
                        flow_steps_map.setdefault(key, []).extend(case.steps or [])
                failed_in_batch = len(batch) - len(cases)
                if failed_in_batch > 0:
                    task.failed += failed_in_batch
                task.progress = (task.completed + task.failed) / max(1, task.total)
                task.logs.append(
                    f"流程批 {idx}/{len(process_batches)} 完成：成功 {len(cases)}，失败 {failed_in_batch}，进度 {task.completed}/{task.total}"
                )

            # 规则批次：按每 20 条分段，可并发请求，谁先返回谁先入库（不影响 JSON 格式）
            if rule_points:
                rule_batches = _chunk_list(rule_points, RULE_BATCH_SIZE if ENABLE_RULE_BATCH else BATCH_SIZE)
                task.logs.append(f"阶段：规则用例生成，共 {len(rule_points)} 条，批次数 {len(rule_batches)}")
            else:
                rule_batches = []
                task.logs.append("阶段：规则用例生成，0 条（跳过）")

            if ENABLE_RULE_BATCH and rule_batches:
                concurrency = max(1, min(RULE_BATCH_CONCURRENCY, len(rule_batches)))
                task.logs.append(f"规则批量请求：已启用，每批 {RULE_BATCH_SIZE}，并发 {concurrency}")
                with ThreadPoolExecutor(max_workers=concurrency, thread_name_prefix="rule_batch") as pool:
                    futures = {
                        pool.submit(
                            self._generator.generate_cases_batch,
                            batch,
                            task.strategy,
                            flow_steps_map,
                            True,
                            task.prompt_examples
                        ): (i, batch)
                        for i, batch in enumerate(rule_batches, start=1)
                    }
                    for future in as_completed(futures):
                        idx, batch = futures[future]
                        cases, tokens, batch_logs = future.result()
                        task.token_usage += tokens
                        task.logs.extend(batch_logs)
                        for case in cases:
                            task.cases.append(case)
                            task.completed += 1
                        failed_in_batch = len(batch) - len(cases)
                        if failed_in_batch > 0:
                            task.failed += failed_in_batch
                        task.progress = (task.completed + task.failed) / max(1, task.total)
                        task.logs.append(
                            f"规则批 {idx}/{len(rule_batches)} 完成：成功 {len(cases)}，失败 {failed_in_batch}，进度 {task.completed}/{task.total}"
                        )
            else:
                for idx, batch in enumerate(rule_batches, start=1):
                    cases, tokens, batch_logs = self._generator.generate_cases_batch(
                        batch,
                        task.strategy,
                        flow_steps_map=flow_steps_map,
                        enable_batch=False,
                        prompt_examples=task.prompt_examples
                    )
                    task.token_usage += tokens
                    task.logs.extend(batch_logs)
                    for case in cases:
                        task.cases.append(case)
                        task.completed += 1
                    failed_in_batch = len(batch) - len(cases)
                    if failed_in_batch > 0:
                        task.failed += failed_in_batch
                    task.progress = (task.completed + task.failed) / max(1, task.total)
                    task.logs.append(
                        f"规则批 {idx}/{len(rule_batches)} 完成：成功 {len(cases)}，失败 {failed_in_batch}，进度 {task.completed}/{task.total}"
                    )

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
