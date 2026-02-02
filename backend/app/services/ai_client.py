"""
百炼模型调用封装
"""
import json
import os
import re
from typing import Any, Tuple

from openai import OpenAI
from app.utils.logger import generator_logger


class AIClient:
    """百炼兼容接口客户端"""

    def __init__(self):
        api_key = os.getenv("DASHSCOPE_API_KEY")
        if not api_key:
            raise ValueError("缺少 DASHSCOPE_API_KEY 环境变量（请在运行环境中配置）")
        self._client = OpenAI(
            api_key=api_key,
            base_url="https://dashscope.aliyuncs.com/compatible-mode/v1"
        )

    def chat_json(
        self,
        system_prompt: str,
        user_prompt: str,
        model: str = "qwen-plus",
        temperature: float = 0.2,
        max_tokens: int = 800
    ) -> Tuple[Any, int]:
        response = self._client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": system_prompt},
                {"role": "user", "content": user_prompt}
            ],
            temperature=temperature,
            max_tokens=max_tokens
        )
        content = response.choices[0].message.content or ""
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
