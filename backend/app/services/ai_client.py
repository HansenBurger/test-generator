"""
百炼模型调用封装
"""
import json
import os
import re
import subprocess
from typing import Any, Optional, Tuple

from openai import OpenAI, APIConnectionError, APIStatusError, APITimeoutError, RateLimitError
from app.core.config import (
    DASHSCOPE_API_BASE_URL,
    DASHSCOPE_DEFAULT_MODEL,
    DASHSCOPE_SUPPORTED_MODELS
)
from app.utils.logger import generator_logger

# 思考模式与 token 余量可通过环境变量配置，便于内网不改代码即可调整：
#   DASHSCOPE_ENABLE_THINKING        是否启用模型思考（true/false，默认 false）
#   DASHSCOPE_THINKING_TOKEN_BUFFER  为思考预留的 token 余量（默认 1500）
# max_tokens 是生成上限而非固定消耗：思考关闭时余量仅放宽上限、不增加实际消耗；
# 思考开启（或网关不透传关闭指令）时，余量保证思考之后仍有空间输出 content。
def _env_bool(name: str, default: bool) -> bool:
    raw = os.getenv(name, "").strip().lower()
    return raw in ("1", "true", "yes", "on") if raw else default


def _env_int(name: str, default: int) -> int:
    raw = os.getenv(name, "").strip()
    if not raw:
        return default
    try:
        return int(raw)
    except ValueError:
        return default


# 默认值（表为空 / 字段为 NULL 时回退到这些，来源 .env 或硬编码）
DEFAULT_ENABLE_THINKING = _env_bool("DASHSCOPE_ENABLE_THINKING", False)
DEFAULT_THINKING_TOKEN_BUFFER = _env_int("DASHSCOPE_THINKING_TOKEN_BUFFER", 1500)

import threading
import time
from app.db import repository as _repo

_runtime_cache = None          # 运行时配置缓存（dict）
_models_cache = None           # {"items": [...], "ts": float}
_ai_client_singleton = None    # AIClient 单例
_state_lock = threading.Lock()
_last_good_model = None        # auto sticky：上次成功的模型
_bad_models = {}               # auto：近期失败模型 -> 时间戳
_BAD_MODEL_TTL = 300
_MODELS_CACHE_TTL = 60
# 探测 /v1/models 时，剔除明显不能用于对话/生成的模型（embedding/音频/图像/视频等），
# 避免下拉里混入误选项。白名单（DASHSCOPE_SUPPORTED_MODELS）不受此过滤影响。
_NON_CHAT_MARKERS = (
    "embed", "audio", "realtime", "tts", "ocr",
    "rerank", "image", "video", "speech", "asr",
)


def _is_chat_model(model_id: str) -> bool:
    mid = (model_id or "").lower()
    return not any(marker in mid for marker in _NON_CHAT_MARKERS)


def _is_failover_error(exc: BaseException) -> bool:
    """判断异常是否属于“应换下一个模型”的类别（额度/限流/不可用/超时）。
    JSON 解析失败等业务异常返回 False，不触发轮转。"""
    if isinstance(exc, (RateLimitError, APITimeoutError, APIConnectionError)):
        return True
    if isinstance(exc, APIStatusError):
        code = getattr(exc, "status_code", None)
        if code in (429, 500, 502, 503, 504):
            return True
        blob = f"{getattr(exc, 'message', '') or ''} {getattr(exc, 'body', '') or ''}".lower()
        if any(k in blob for k in (
            "quota", "insufficient", "arrearage", "balance",
            "throttl", "limit", "额度", "不存在", "not found", "not activated", "activated",
        )):
            return True
        return False
    return False


def _err_brief(exc: BaseException) -> str:
    return str(exc)[:160].replace("\n", " ")


def _default_config() -> dict:
    return {
        "enable_thinking": DEFAULT_ENABLE_THINKING,
        "thinking_token_buffer": DEFAULT_THINKING_TOKEN_BUFFER,
        "current_model": None,
        "temperature": None,
        "model_mode": "default",
    }


def _merge_config(record) -> dict:
    base = _default_config()
    if record is None:
        return base
    if record.enable_thinking is not None:
        base["enable_thinking"] = bool(record.enable_thinking)
    if record.thinking_token_buffer is not None:
        base["thinking_token_buffer"] = int(record.thinking_token_buffer)
    base["current_model"] = record.current_model or None
    base["temperature"] = float(record.temperature) if record.temperature is not None else None
    mm = record.model_mode
    base["model_mode"] = mm if mm in ("fixed", "auto", "default") else "default"
    return base


def get_runtime_config(force: bool = False) -> dict:
    """读取运行时配置：进程内缓存优先，否则查 DB 并与默认值合并。"""
    global _runtime_cache
    with _state_lock:
        if _runtime_cache is not None and not force:
            return dict(_runtime_cache)
    merged = _merge_config(_repo.get_model_config())
    with _state_lock:
        _runtime_cache = merged
    return dict(merged)


def apply_model_config(
    enable_thinking: bool,
    thinking_token_buffer: int,
    current_model: Optional[str],
    temperature: Optional[float] = None,
    model_mode: Optional[str] = None,
) -> dict:
    """写入运行时配置到 DB 并刷新缓存，返回合并后的配置。"""
    global _runtime_cache
    norm_mode = model_mode if model_mode in ("fixed", "auto", "default") else "default"
    eff_current = (current_model or None) if norm_mode == "fixed" else None
    eff_temp = float(temperature) if temperature is not None else None
    record = _repo.upsert_model_config(
        enable_thinking=bool(enable_thinking),
        thinking_token_buffer=int(thinking_token_buffer),
        current_model=eff_current,
        temperature=eff_temp,
        model_mode=norm_mode,
    )
    merged = _merge_config(record)
    with _state_lock:
        _runtime_cache = merged
    return merged


class AIClient:
    """百炼兼容接口客户端"""

    def __init__(self):
        api_key = os.getenv("DASHSCOPE_API_KEY")
        source = "process"
        if not api_key and os.name == "nt":
            api_key = self._read_windows_env("DASHSCOPE_API_KEY")
            source = "windows"
        if not api_key:
            raise ValueError("缺少 DASHSCOPE_API_KEY 环境变量（请在运行环境中配置）")
        masked = f"{api_key[:4]}...{api_key[-4:]}" if len(api_key) >= 8 else "***"
        generator_logger.info(
            "DASHSCOPE_API_KEY loaded from %s (length=%s, masked=%s)",
            source,
            len(api_key),
            masked,
        )
        self._client = OpenAI(
            api_key=api_key,
            base_url=DASHSCOPE_API_BASE_URL,
        )

    def chat_json(
        self,
        system_prompt: str,
        user_prompt: str,
        model: Optional[str] = None,
        temperature: float = 0.2,
        max_tokens: int = 800
    ) -> Tuple[Any, int]:
        cfg = get_runtime_config()
        mode = cfg.get("model_mode") or "default"
        explicit = (model or "").strip()
        # 预留思考余量，兼容 Qwen3 等默认 thinking 的模型，防止 content 为空。
        effective_max_tokens = max_tokens + int(cfg["thinking_token_buffer"])
        extra_body = {}
        if not cfg["enable_thinking"]:
            extra_body["enable_thinking"] = False
            extra_body["chat_template_kwargs"] = {"enable_thinking": False}

        def _check_supported(m: str) -> None:
            if DASHSCOPE_SUPPORTED_MODELS and m not in DASHSCOPE_SUPPORTED_MODELS:
                raise ValueError(f"不支持的模型: {m}")

        # 1) 调用方显式指定模型：最高优先，单模型，不轮转
        if explicit:
            _check_supported(explicit)
            return self._invoke_once(
                explicit, system_prompt, user_prompt, temperature, effective_max_tokens, extra_body
            )

        # 2) auto：在可选模型列表内 sticky + 失败轮转（“用完/失败一个换下一个”）
        if mode == "auto":
            ordered = self._order_candidates(self._auto_candidates())
            last_err: Optional[Exception] = None
            for m in ordered:
                try:
                    result = self._invoke_once(
                        m, system_prompt, user_prompt, temperature, effective_max_tokens, extra_body
                    )
                    self._set_good_model(m)
                    return result
                except Exception as exc:
                    if _is_failover_error(exc):
                        generator_logger.warning(
                            "auto 轮转：模型 %s 不可用（%s），尝试下一个", m, _err_brief(exc)
                        )
                        self._mark_bad_model(m)
                        last_err = exc
                        continue
                    raise
            if last_err is not None:
                raise last_err
            raise RuntimeError("auto 模式无可用模型")

        # 3) fixed / default
        if mode == "fixed" and cfg["current_model"]:
            selected = cfg["current_model"]
        else:
            selected = DASHSCOPE_DEFAULT_MODEL
        _check_supported(selected)
        return self._invoke_once(
            selected, system_prompt, user_prompt, temperature, effective_max_tokens, extra_body
        )

    def _invoke_once(
        self,
        model: str,
        system_prompt: str,
        user_prompt: str,
        temperature: float,
        effective_max_tokens: int,
        extra_body: dict,
    ) -> Tuple[Any, int]:
        """对单个模型发起一次调用并解析 JSON；失败原样抛出（轮转判断由调用方负责）。"""
        response = self._client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=temperature,
            max_tokens=effective_max_tokens,
            extra_body=extra_body,
        )
        message = response.choices[0].message
        content = message.content or ""
        if not content:
            # Qwen3 等模型可能将内容放在 reasoning_content 中
            reasoning = getattr(message, "reasoning_content", None) or ""
            generator_logger.warning(
                "模型返回 content 为空，reasoning_content 长度=%d，完整 message=%s",
                len(reasoning),
                message.model_dump_json() if hasattr(message, "model_dump_json") else str(message),
            )
        try:
            json_data = self._extract_json(content)
        except json.JSONDecodeError as exc:
            # 这里的错误基本都是“模型输出不合法 JSON”（常见：输出被截断、缺少逗号/括号、混入解释文本）
            head = content[:400].replace("\n", "\\n")
            tail = content[-400:].replace("\n", "\\n") if len(content) > 400 else ""
            generator_logger.warning(
                "模型返回 JSON 解析失败：%s；content_head=%s%s",
                str(exc),
                head,
                f"；content_tail={tail}" if tail else "",
            )
            raise ValueError(f"模型返回 JSON 不合法：{str(exc)}") from exc
        usage = getattr(response, "usage", None)
        total_tokens = int(getattr(usage, "total_tokens", 0) or 0)
        return json_data, total_tokens

    def _auto_candidates(self) -> list:
        """auto 轮转候选模型：复用可选模型列表（白名单优先，否则探测过滤后列表）。"""
        try:
            cand = self.list_available_models()
        except Exception as exc:
            generator_logger.warning("auto 获取候选模型失败：%s", exc)
            cand = []
        seen = set()
        uniq = []
        for m in cand:
            if m and m not in seen:
                seen.add(m)
                uniq.append(m)
        return uniq or [DASHSCOPE_DEFAULT_MODEL]

    def _order_candidates(self, candidates: list) -> list:
        """sticky 排序：上次成功模型优先，近期失败模型沉底，其余保持原序。"""
        now = time.time()
        with _state_lock:
            good = _last_good_model
            bad = {m for m, t in _bad_models.items() if now - t < _BAD_MODEL_TTL}

        def rank(m: str) -> int:
            if m == good:
                return 0
            if m in bad:
                return 2
            return 1

        return sorted(candidates, key=lambda m: (rank(m), candidates.index(m)))

    def _mark_bad_model(self, model: str) -> None:
        with _state_lock:
            _bad_models[model] = time.time()

    def _set_good_model(self, model: str) -> None:
        global _last_good_model
        with _state_lock:
            _last_good_model = model
            _bad_models.pop(model, None)

    def _read_windows_env(self, name: str) -> str:
        try:
            result = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-NonInteractive",
                    "-Command",
                    "[Environment]::GetEnvironmentVariable($args[0], 'User')",
                    name,
                ],
                capture_output=True,
                text=True,
                timeout=5,
                check=False,
            )
            value = (result.stdout or "").strip()
            if value:
                return value
            result = subprocess.run(
                [
                    "powershell",
                    "-NoProfile",
                    "-NonInteractive",
                    "-Command",
                    "[Environment]::GetEnvironmentVariable($args[0], 'Machine')",
                    name,
                ],
                capture_output=True,
                text=True,
                timeout=5,
                check=False,
            )
            return (result.stdout or "").strip()
        except Exception:
            return ""

    def _extract_json(self, content: str) -> Any:
        content = content.strip()
        if not content:
            raise ValueError("模型返回内容为空")

        # 优先解析代码块内容
        if "```" in content:
            fence_match = re.search(r"```(?:json)?\s*([\s\S]*?)\s*```", content, re.IGNORECASE)
            if fence_match:
                content = fence_match.group(1).strip()

        # 从首个 JSON 起始符开始解析，忽略尾部多余内容
        start_candidates = [pos for pos in (content.find("{"), content.find("[")) if pos != -1]
        if not start_candidates:
            raise ValueError("模型返回内容无法解析为JSON")
        start = min(start_candidates)
        decoder = json.JSONDecoder()
        try:
            obj, _ = decoder.raw_decode(content[start:])
            return obj
        except json.JSONDecodeError:
            trimmed = self._trim_json_tail(content[start:])
            if trimmed:
                obj, _ = decoder.raw_decode(trimmed)
                return obj
            raise

    def _trim_json_tail(self, text: str) -> str:
        last_obj = text.rfind("}")
        last_arr = text.rfind("]")
        last = max(last_obj, last_arr)
        if last == -1:
            return ""
        return text[: last + 1]

    def list_available_models(self) -> list:
        """可选模型列表：白名单优先，否则实时探测网关 /models，失败回退默认模型。"""
        if DASHSCOPE_SUPPORTED_MODELS:
            return list(DASHSCOPE_SUPPORTED_MODELS)
        global _models_cache
        now = time.time()
        with _state_lock:
            if _models_cache and (now - _models_cache["ts"]) < _MODELS_CACHE_TTL:
                return list(_models_cache["items"])
        items = []
        try:
            page = self._client.models.list()
            skipped = 0
            for m in getattr(page, "data", []) or []:
                mid = getattr(m, "id", None)
                if not mid:
                    continue
                if _is_chat_model(mid):
                    items.append(mid)
                else:
                    skipped += 1
            if skipped:
                generator_logger.info(
                    "模型列表过滤掉 %d 个非对话模型（embedding/音频/图像等）", skipped
                )
        except Exception as exc:
            generator_logger.warning("探测可用模型列表失败，回退默认模型：%s", exc)
        if not items:
            items = [DASHSCOPE_DEFAULT_MODEL]
        with _state_lock:
            _models_cache = {"items": items, "ts": now}
        return items


def get_ai_client() -> "AIClient":
    """AIClient 懒加载单例。"""
    global _ai_client_singleton
    if _ai_client_singleton is None:
        with _state_lock:
            if _ai_client_singleton is None:
                _ai_client_singleton = AIClient()
    return _ai_client_singleton


def get_config_view() -> dict:
    """组装前端配置面板所需的完整视图。"""
    cfg = get_runtime_config()
    try:
        available = get_ai_client().list_available_models()
    except Exception as exc:
        generator_logger.warning("获取可用模型列表失败：%s", exc)
        available = [DASHSCOPE_DEFAULT_MODEL]
    mode = cfg.get("model_mode") or "default"
    if mode == "auto":
        effective_model = "auto"
    elif mode == "fixed" and cfg["current_model"]:
        effective_model = cfg["current_model"]
    else:
        effective_model = DASHSCOPE_DEFAULT_MODEL
    return {
        "enable_thinking": cfg["enable_thinking"],
        "thinking_token_buffer": cfg["thinking_token_buffer"],
        "current_model": cfg["current_model"],
        "temperature": cfg.get("temperature"),
        "model_mode": mode,
        "effective_model": effective_model,
        "available_models": available,
        "default_model": DASHSCOPE_DEFAULT_MODEL,
        "defaults": _default_config(),
    }
