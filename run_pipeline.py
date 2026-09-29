#!/usr/bin/env python3
"""运行 PPT Parsing Pipeline 解析单个 pptx 文件的脚本。

用法：
  python src/run_pipeline.py input.pptx [output.md]
  python src/run_pipeline.py input.pptx --output output.md --verbose
"""

import argparse
import sys
from pathlib import Path

# 保证能 import src 包
sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

from src.pipeline import PPTParsingPipeline


def main():
    parser = argparse.ArgumentParser(description="解析 PPTX 文件为 Markdown")
    parser.add_argument("input", help="输入 PPTX 文件路径")
    parser.add_argument("output", nargs="?", default=None,
                        help="输出 Markdown 文件路径（默认同名 .md）")
    parser.add_argument("--verbose", "-v", action="store_true", help="输出详细日志")
    parser.add_argument("--no-validation", action="store_true", help="禁用 Validation Agent")
    parser.add_argument("--no-reorder", action="store_true", help="禁用 VLM 阅读顺序重排")
    parser.add_argument("--model", default=None, help="指定 VLM 模型名称")
    args = parser.parse_args()

    input_path = Path(args.input)
    if not input_path.exists():
        print(f"错误：文件不存在 {input_path}")
        sys.exit(1)

    if args.output is None:
        output_path = input_path.with_suffix(".md")
    else:
        output_path = Path(args.output)

    from src.llm.qwen_multimodal import QwenMultimodalClient

    vlm_client = QwenMultimodalClient(model=args.model) if args.model else None

    pipeline = PPTParsingPipeline(
        vlm_client=vlm_client,
        enable_validation=not args.no_validation,
        enable_reorder=not args.no_reorder,
        verbose=args.verbose,
    )

    md = pipeline.parse(str(input_path), str(output_path))
    print(f"\n解析完成，输出：{output_path}")
    print(f"Markdown 长度：{len(md)} 字符")


if __name__ == "__main__":
    main()
