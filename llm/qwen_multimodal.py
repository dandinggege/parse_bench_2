#!/usr/bin/env python3
"""通义千问多模态大模型调用客户端。

支持：
  - 纯文本对话
  - 单图/多图输入（本地文件路径 或 图片 URL）
  - 深度思考模式（enable_thinking）

用法示例：
  from src.llm.qwen_multimodal import QwenMultimodalClient

  client = QwenMultimodalClient()

  # 纯文本
  result = client.chat("你好")

  # 单张本地图片
  result = client.chat("描述这张图片", images=["/path/to/image.png"])

  # 图片 URL
  result = client.chat("描述这张图片", images=["https://example.com/img.jpg"])

  # 多图
  result = client.chat("比较这两张图", images=["a.png", "b.png"])
"""

import base64
import mimetypes
import os
from pathlib import Path

from openai import OpenAI


DEFAULT_BASE_URL = "https://llm-snr3c33splncmm6m.cn-beijing.maas.aliyuncs.com/compatible-mode/v1"
DEFAULT_MODEL = "qwen-vl-max"  # 多模态视觉模型
DEFAULT_API_KEY = "sk-ws-H.PLXXRMH.EOi8.MEUCIBkBKc12_cVxbY4QNNZZT_ICxGOx0-aZ5fb1hlX5T8AqAiEAkp29UPD3jUdDjcql7Qtm0BEFN9cjiCM68kP9tpbAx_0"


class QwenMultimodalClient:
    """通义千问多模态调用客户端。"""

    def __init__(
        self,
        api_key: str | None = None,
        base_url: str = DEFAULT_BASE_URL,
        model: str = DEFAULT_MODEL,
        enable_thinking: bool = False,
    ):
        self.model = model
        self.enable_thinking = enable_thinking
        self._client = OpenAI(
            api_key=api_key or os.getenv("DASHSCOPE_API_KEY") or DEFAULT_API_KEY,
            base_url=base_url,
        )

    def chat(
        self,
        prompt: str,
        images: list[str] | None = None,
        system_prompt: str | None = None,
        max_tokens: int = 4096,
    ) -> dict:
        """发送多模态请求，返回结果。

        Args:
            prompt: 用户文本提示
            images: 图片列表，每项可以是本地文件路径或 URL
            system_prompt: 系统提示（可选）
            max_tokens: 最大输出 token 数

        Returns:
            dict: {"content": str, "reasoning": str | None, "usage": dict}
        """
        messages = []
        if system_prompt:
            messages.append({"role": "system", "content": system_prompt})

        user_content = self._build_user_content(prompt, images)
        messages.append({"role": "user", "content": user_content})

        extra_body = {}
        if self.enable_thinking:
            extra_body["enable_thinking"] = True

        completion = self._client.chat.completions.create(
            model=self.model,
            messages=messages,
            max_tokens=max_tokens,
            extra_body=extra_body or None,
            stream=True,
        )

        return self._collect_stream(completion)

    def _build_user_content(self, prompt: str, images: list[str] | None) -> list[dict]:
        """构建 OpenAI vision API 格式的 user content。"""
        parts = []

        if images:
            for img in images:
                url = self._to_image_url(img)
                parts.append({
                    "type": "image_url",
                    "image_url": {"url": url},
                })

        parts.append({"type": "text", "text": prompt})
        return parts

    def _to_image_url(self, image: str) -> str:
        """将图片路径或 URL 转为 API 可用的 URL。

        - 已是 URL（http/https）：直接返回
        - 已是 data URI（data:...;base64,...）：直接返回
        - 本地文件路径：读取并转 base64 data URI
        """
        if image.startswith(("http://", "https://", "data:")):
            return image

        path = Path(image)
        if not path.exists():
            raise FileNotFoundError(f"图片文件不存在: {image}")

        mime = mimetypes.guess_type(str(path))[0] or "image/png"
        with open(path, "rb") as f:
            data = base64.b64encode(f.read()).decode("utf-8")
        return f"data:{mime};base64,{data}"

    def _collect_stream(self, completion) -> dict:
        """收集流式响应，拼接 content 和 reasoning_content。"""
        content_parts = []
        reasoning_parts = []

        for chunk in completion:
            if not chunk.choices:
                continue
            delta = chunk.choices[0].delta

            # 收集推理内容（深度思考模式）
            if hasattr(delta, "reasoning_content") and delta.reasoning_content:
                reasoning_parts.append(delta.reasoning_content)

            # 收集正文内容
            if delta.content:
                content_parts.append(delta.content)

        result = {
            "content": "".join(content_parts),
            "reasoning": "".join(reasoning_parts) if reasoning_parts else None,
            "usage": {},
        }

        # 最后一个 chunk 通常带 usage
        try:
            if chunk.usage:
                result["usage"] = {
                    "prompt_tokens": chunk.usage.prompt_tokens,
                    "completion_tokens": chunk.usage.completion_tokens,
                    "total_tokens": chunk.usage.total_tokens,
                }
        except (AttributeError, NameError):
            pass

        return result


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="通义千问多模态调用测试")
    parser.add_argument("prompt", help="文本提示")
    parser.add_argument("--images", nargs="*", default=[], help="图片路径或 URL")
    parser.add_argument("--model", default=DEFAULT_MODEL, help=f"模型名称 (默认: {DEFAULT_MODEL})")
    parser.add_argument("--thinking", action="store_true", help="启用深度思考模式")
    parser.add_argument("--system", default=None, help="系统提示")
    args = parser.parse_args()

    client = QwenMultimodalClient(model=args.model, enable_thinking=args.thinking)
    result = client.chat(args.prompt, images=args.images, system_prompt=args.system)

    if result["reasoning"]:
        print("=" * 20 + " 思考过程 " + "=" * 20)
        print(result["reasoning"])
        print("=" * 50)

    print("\n" + result["content"])

    if result["usage"]:
        print(f"\n[Token 用量] prompt: {result['usage'].get('prompt_tokens', '?')}, "
              f"completion: {result['usage'].get('completion_tokens', '?')}, "
              f"total: {result['usage'].get('total_tokens', '?')}")
