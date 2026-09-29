#!/usr/bin/env python3
"""PPT Parsing Pipeline - 主流程。

流程：
1. File Analyzer → 双路解析（Native + VLM）
2. Slide Understanding → 单页理解
3. Layout Understanding → 布局分析
4. Semantic Understanding → Agent 处理
5. Slide Reconstruction → 重构 Markdown
6. Validation → 质量校验（可回退重处理）
"""

import tempfile
from pathlib import Path

from ..llm.qwen_multimodal import QwenMultimodalClient
from .file_analyzer import FileAnalyzer, SlideData
from .slide_renderer import SlideRenderer
from .agents import TextAgent, TableAgent, VisionAgent, ValidationAgent


class PPTParsingPipeline:
    """PPT 解析主流程。"""

    def __init__(
        self,
        vlm_client: QwenMultimodalClient | None = None,
        enable_validation: bool = True,
        enable_reorder: bool = True,
        max_retries: int = 1,
        verbose: bool = False,
    ):
        self.vlm_client = vlm_client or QwenMultimodalClient()
        self.enable_validation = enable_validation
        self.enable_reorder = enable_reorder
        self.max_retries = max_retries
        self.verbose = verbose

        # 初始化 Agents
        self.text_agent = TextAgent()
        self.table_agent = TableAgent()
        self.vision_agent = VisionAgent(self.vlm_client)
        self.validation_agent = ValidationAgent(self.vlm_client)

    def parse(self, pptx_path: str, output_path: str | None = None) -> str:
        """解析 PPTX 文件，返回 Markdown。

        Args:
            pptx_path: PPTX 文件路径
            output_path: 输出 Markdown 文件路径（可选）

        Returns:
            str: 完整的 Markdown 文本
        """
        pptx_path = Path(pptx_path)
        if not pptx_path.exists():
            raise FileNotFoundError(f"文件不存在: {pptx_path}")

        self._log(f"开始解析: {pptx_path.name}")

        # ① File Analyzer - 双路解析
        self._log("① File Analyzer - Native Parsing...")
        analyzer = FileAnalyzer(str(pptx_path))
        slides_data = analyzer.analyze_all_slides()
        self._log(f"  Native 解析完成: {len(slides_data)} 页")

        # 渲染幻灯片为图片（供 VLM 使用）
        self._log("① File Analyzer - Slide Rendering...")
        renderer = SlideRenderer(str(pptx_path))
        slide_images = renderer.render_all_slides()
        self._log(f"  渲染完成: {len(slide_images)} 张图片")

        # 逐页处理
        all_markdown = []
        for slide_data in slides_data:
            slide_idx = slide_data.slide_index
            slide_image = slide_images[slide_idx] if slide_idx < len(slide_images) else None

            self._log(f"\n处理第 {slide_idx + 1} 页...")

            # 处理（可重试）
            for attempt in range(self.max_retries + 1):
                if attempt > 0:
                    self._log(f"  重试 (第 {attempt} 次)...")

                md = self._process_slide(slide_data, slide_image)

                # Validation
                if self.enable_validation and slide_image:
                    self._log("  ⑥ Validation...")
                    validation = self.validation_agent.validate(slide_data, md, slide_image)
                    self._log(f"    分数: {validation['score']}, 通过: {validation['pass']}")

                    if validation["pass"] or attempt == self.max_retries:
                        all_markdown.append(md)
                        break
                else:
                    all_markdown.append(md)
                    break

        # 合并结果
        full_markdown = "\n\n---\n\n".join(all_markdown)

        # 写入文件
        if output_path:
            output_path = Path(output_path)
            output_path.parent.mkdir(parents=True, exist_ok=True)
            output_path.write_text(full_markdown, encoding="utf-8")
            self._log(f"\n已保存到: {output_path}")

        self._log(f"\n解析完成: {len(slides_data)} 页 → {len(all_markdown)} 个 Markdown 段落")
        return full_markdown

    def _process_slide(self, slide_data: SlideData, slide_image: Path | None) -> str:
        """处理单页幻灯片。"""
        md_parts = []

        # ② Slide Understanding - 分类元素
        self._log("  ② Slide Understanding...")

        # 标题放最前
        if slide_data.title:
            md_parts.append(f"# {slide_data.title}\n")

        # ③ Layout Understanding - 按页面位置排序所有元素
        self._log("  ③ Layout Understanding...")
        sorted_elements = self._sort_by_layout(slide_data)

        # ④ Semantic Understanding + ⑤ Slide Reconstruction
        # 按阅读顺序（上到下、左到右）逐个元素转换输出
        self._log("  ④ Semantic Understanding + ⑤ Slide Reconstruction...")
        for etype, element in sorted_elements:
            md = self._element_to_markdown(etype, element, slide_data.title)
            if md:
                md_parts.append(md)

        raw_md = "\n".join(md_parts)

        # ⑤ VLM 视觉重排 - 结合页面图片按阅读顺序重新排版
        if self.enable_reorder and slide_image:
            self._log("  ⑤ VLM 阅读顺序重排...")
            reordered = self._reorder_with_vlm(slide_data, raw_md, slide_image)
            if reordered:
                return reordered

        return raw_md

    def _reorder_with_vlm(self, slide_data: SlideData, raw_md: str, slide_image: Path) -> str | None:
        """调用 VLM，结合页面图片按阅读顺序重新排版内容。"""
        prompt = f"""You are a document layout expert. Below is a rendered image of a PPT slide and the parsed content extracted from it. The content order may not match the visual reading order.

Please reorder and reformat the content into well-structured Markdown according to the visual reading order in the image (top-to-bottom, left-to-right; for multi-column layouts, read the left column first, then the right column).

Rules:
1. Do NOT rewrite, add, or fabricate any content — only reorder and reformat what is given.
2. Preserve tables as HTML <table>, formulas as LaTeX.
3. Use appropriate heading levels (#, ##, ###).
4. Do NOT generate any image URLs or links (no https://, no markdown image syntax). If an image is referenced, keep only its plain-text caption.
5. Output only the Markdown, with no extra explanation.

[Parsed content]
{raw_md}"""

        try:
            result = self.vlm_client.chat(prompt, images=[str(slide_image)])
            content = result.get("content", "").strip()
            if content:
                return content
        except Exception as e:
            self._log(f"    VLM 重排失败: {e}")
        return None

    def _element_to_markdown(self, etype: str, element, title: str) -> str:
        """将单个元素转为 Markdown。"""
        if etype == "text":
            level = self._heading_level(element.font_size)
            if level > 0:
                return f"\n{'#' * level} {element.text}\n"
            return f"\n{element.text}\n"

        if etype == "table":
            if element.html:
                return f"\n{element.html}\n"
            return self.table_agent._simple_table_html(element)

        if etype == "chart":
            self._log("    Vision Agent 处理图表...")
            return self.vision_agent.process_chart(element.chart_type, element.data)

        if etype == "image":
            self._log("    Vision Agent 处理图片...")
            descs = self.vision_agent.process_images([element], context=f"Slide title: {title}")
            if descs and descs[0]:
                return f"\n[Image: {descs[0]}]\n"
            return ""

        if etype == "shape":
            if element.text:
                return f"\n{element.text}\n"
            return ""

        return ""

    @staticmethod
    def _heading_level(font_size: float) -> int:
        """根据字号判断标题层级。"""
        if font_size >= 40:
            return 1
        if font_size >= 26:
            return 2
        if font_size >= 20:
            return 3
        return 0

    def _sort_by_layout(self, slide_data: SlideData) -> list:
        """按布局位置排序元素（从上到下，从左到右），返回 (类型, 元素) 列表。"""
        elements = []
        for text_el in slide_data.texts:
            if text_el.is_title:
                continue  # 标题已单独处理
            elements.append(("text", text_el, text_el.bbox))
        for tbl_el in slide_data.tables:
            elements.append(("table", tbl_el, tbl_el.bbox))
        for chart_el in slide_data.charts:
            elements.append(("chart", chart_el, chart_el.bbox))
        for img_el in slide_data.images:
            elements.append(("image", img_el, img_el.bbox))
        for shape_el in slide_data.shapes:
            elements.append(("shape", shape_el, shape_el.bbox))

        # 按 y1 排序（从上到下），再按 x1（从左到右）
        elements.sort(key=lambda x: (x[2][1], x[2][0]))
        return [(e[0], e[1]) for e in elements]

    def _log(self, msg: str):
        """输出日志。"""
        if self.verbose:
            print(msg)


def parse_pptx(
    pptx_path: str,
    output_path: str | None = None,
    verbose: bool = False,
) -> str:
    """便捷函数：解析 PPTX 文件。

    Args:
        pptx_path: PPTX 文件路径
        output_path: 输出 Markdown 文件路径（可选）
        verbose: 是否输出详细日志

    Returns:
        str: Markdown 文本
    """
    pipeline = PPTParsingPipeline(verbose=verbose)
    return pipeline.parse(pptx_path, output_path)


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="PPT Parsing Pipeline")
    parser.add_argument("input", help="输入 PPTX 文件路径")
    parser.add_argument("output", nargs="?", help="输出 Markdown 文件路径")
    parser.add_argument("--verbose", "-v", action="store_true", help="输出详细日志")
    parser.add_argument("--no-validation", action="store_true", help="禁用验证")
    args = parser.parse_args()

    pipeline = PPTParsingPipeline(
        verbose=args.verbose,
        enable_validation=not args.no_validation,
    )
    pipeline.parse(args.input, args.output)
