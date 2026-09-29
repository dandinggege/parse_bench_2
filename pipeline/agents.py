#!/usr/bin/env python3
"""Agents - 各种处理 Agent。

- TextAgent: 处理文本内容
- TableAgent: 处理表格
- VisionAgent: 处理图片/图表（调用 VLM）
- ValidationAgent: 质量校验
"""

import base64
import json
from pathlib import Path

from ..llm.qwen_multimodal import QwenMultimodalClient
from .file_analyzer import SlideData, TextElement, TableElement, ImageElement


class TextAgent:
    """文本处理 Agent - 将文本元素转为 Markdown。"""

    def __init__(self):
        pass

    def process(self, texts: list[TextElement], title: str = "") -> str:
        """处理文本元素列表，返回 Markdown。"""
        md_parts = []

        # 添加标题
        if title:
            md_parts.append(f"# {title}\n")

        # 处理其他文本
        for text_el in texts:
            if text_el.is_title:
                continue  # 标题已处理

            # 根据字号判断标题层级
            level = self._detect_heading_level(text_el.font_size)
            if level > 0:
                md_parts.append(f"\n{'#' * level} {text_el.text}\n")
            else:
                # 普通文本
                md_parts.append(f"\n{text_el.text}\n")

        return "\n".join(md_parts)

    def _detect_heading_level(self, font_size: float) -> int:
        """根据字号判断标题层级。"""
        if font_size >= 40:
            return 1
        if font_size >= 26:
            return 2
        if font_size >= 20:
            return 3
        return 0


class TableAgent:
    """表格处理 Agent - 输出 HTML 表格。"""

    def __init__(self):
        pass

    def process(self, tables: list[TableElement]) -> list[str]:
        """处理表格列表，返回 HTML 列表。"""
        html_parts = []
        for tbl in tables:
            if tbl.html:
                html_parts.append(f"\n{tbl.html}\n")
            else:
                # 回退：简单输出
                html_parts.append(self._simple_table_html(tbl))
        return html_parts

    def _simple_table_html(self, tbl: TableElement) -> str:
        """生成简单 HTML 表格。"""
        html = "<table>\n"
        for i, row in enumerate(tbl.cells):
            html += "  <tr>\n"
            for cell in row:
                tag = "th" if i == 0 else "td"
                html += f"    <{tag}>{cell}</{tag}>\n"
            html += "  </tr>\n"
        html += "</table>"
        return html


class VisionAgent:
    """视觉处理 Agent - 调用 VLM 处理图片/图表。"""

    def __init__(self, vlm_client: QwenMultimodalClient | None = None):
        self.vlm_client = vlm_client or QwenMultimodalClient()

    def process_images(self, images: list[ImageElement], context: str = "") -> list[str]:
        """处理图片列表，返回描述文本。"""
        descriptions = []
        for img in images:
            desc = self._describe_image(img, context)
            if desc:
                descriptions.append(desc)
        return descriptions

    def _describe_image(self, image: ImageElement, context: str) -> str:
        """描述单张图片。"""
        # 将图片字节转为 base64 data URI
        b64 = base64.b64encode(image.image_bytes).decode("utf-8")
        data_uri = f"data:image/png;base64,{b64}"

        prompt = "Please describe the content of this image, including the text, chart type, and data. Answer concisely in English."
        if context:
            prompt += f"\n上下文：{context}"

        try:
            result = self.vlm_client.chat(prompt, images=[data_uri])
            return result["content"]
        except Exception as e:
            err = str(e)
            if len(err) > 200:
                err = err[:200] + "..."
            return f"[图片描述失败: {err}]"

    def process_chart(self, chart_type: str, data: dict) -> str:
        """处理图表数据，生成描述。"""
        prompt = f"""这是一个 {chart_type} 类型的图表。
数据：{json.dumps(data, ensure_ascii=False, indent=2)}

请生成一个 Markdown 表格来展示这些数据，并添加一个标题说明图表类型。"""

        try:
            result = self.vlm_client.chat(prompt)
            return result["content"]
        except Exception as e:
            return f"[图表处理失败: {e}]"


class ValidationAgent:
    """质量校验 Agent - 判断解析结果是否合格。"""

    def __init__(self, vlm_client: QwenMultimodalClient | None = None):
        self.vlm_client = vlm_client or QwenMultimodalClient()

    def validate(
        self,
        slide_data: SlideData,
        markdown_output: str,
        slide_image: Path | None = None,
    ) -> dict:
        """验证解析结果。

        Returns:
            dict: {"pass": bool, "score": float, "issues": list[str], "suggestion": str}
        """
        issues = []

        # 规则校验
        issues.extend(self._check_text_coverage(slide_data, markdown_output))
        issues.extend(self._check_table_completeness(slide_data, markdown_output))

        # 如果有图片，用 VLM 做视觉对比
        vlm_issues = []
        if slide_image:
            vlm_issues = self._vlm_validation(slide_image, markdown_output)
            issues.extend(vlm_issues)

        # 计算分数
        score = max(0, 100 - len(issues) * 10)
        passed = score >= 70 and len(issues) <= 3

        return {
            "pass": passed,
            "score": score,
            "issues": issues,
            "suggestion": "重处理" if not passed else "合格",
        }

    def _check_text_coverage(self, slide_data: SlideData, markdown: str) -> list[str]:
        """检查文本覆盖率。"""
        issues = []
        for text_el in slide_data.texts:
            # 简单检查：文本是否出现在输出中
            # 实际应该用更复杂的相似度算法
            if len(text_el.text) > 10 and text_el.text[:20] not in markdown:
                issues.append(f"文本可能遗漏: {text_el.text[:30]}...")
        return issues

    def _check_table_completeness(self, slide_data: SlideData, markdown: str) -> list[str]:
        """检查表格完整性。"""
        issues = []
        for tbl in slide_data.tables:
            # 检查是否有表格 HTML
            if "<table>" not in markdown:
                issues.append("表格未输出")
                break
            # 检查行数
            if markdown.count("<tr>") < tbl.rows:
                issues.append(f"表格行数不足: 期望 {tbl.rows}，实际可能更少")
        return issues

    def _vlm_validation(self, slide_image: Path, markdown: str) -> list[str]:
        """用 VLM 验证解析质量。"""
        prompt = """你是一个文档解析质量评估专家。

请对比以下两项：
1. 原始幻灯片图片
2. 解析后的 Markdown 文本

评估解析质量，指出明显问题（如果有）：
- 文本遗漏
- 表格结构错误
- 图片/图表未处理
- 格式问题

如果没有问题，回答"无明显问题"。
只列出问题，不要解释。"""

        try:
            result = self.vlm_client.chat(
                prompt,
                images=[str(slide_image)],
            )
            content = result["content"]
            if "无明显问题" in content:
                return []
            # 按行拆分问题
            issues = [line.strip() for line in content.split("\n") if line.strip() and not line.startswith("评估")]
            return issues[:5]  # 最多 5 个问题
        except Exception as e:
            return [f"VLM 验证失败: {e}"]
