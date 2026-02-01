"""
百炼模型调用封装
"""
import json
import os
import re
from typing import Any, Dict, Tuple

from openai import OpenAI


class AIClient:
    """百炼兼容接口客户端"""

    def __init__(self):
        api_key = os.getenv("DASHSCOPE_API_KEY") or "YOUR_DASHSCOPE_API_KEY"
        if not api_key:
            raise ValueError("缺少 DASHSCOPE_API_KEY 环境变量")
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
    ) -> Tuple[Dict[str, Any], int]:
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
        json_data = self._extract_json(content)
        usage = getattr(response, "usage", None)
        total_tokens = int(getattr(usage, "total_tokens", 0) or 0)
        return json_data, total_tokens

    def _extract_json(self, content: str) -> Dict[str, Any]:
        content = content.strip()
        if content.startswith("{") and content.endswith("}"):
            return json.loads(content)

        # 尝试从代码块中提取
        match = re.search(r"\{[\s\S]*\}", content)
        if match:
            return json.loads(match.group(0))

        raise ValueError("模型返回内容无法解析为JSON")
