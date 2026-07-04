#!/usr/bin/env python3
"""
analyze_template.py
============================================================
一步到位：解析 PPT 模板并检测所有 run 的 template/variable 属性。
整合了 parse_template.py 和 detect_template_slots.py 的功能。

Usage:
    python analyze_template.py <template.pptx> <output.json> [--mode heuristic|ai-hint]
"""
import json
import sys
from pathlib import Path

# 引入同目录下的两个脚本
import parse_template as pt
import detect_template_slots as dts


def main():
    if len(sys.argv) < 3:
        print("Usage: python analyze_template.py <template.pptx> <output.json> [--mode heuristic|ai-hint]")
        sys.exit(1)

    pptx_path = sys.argv[1]
    output_path = sys.argv[2]

    mode = "heuristic"
    if "--mode" in sys.argv:
        idx = sys.argv.index("--mode")
        mode = sys.argv[idx + 1]

    # Step 1: 解析模板（版式、字体、配色、装饰）
    template_meta = pt.parse_template(pptx_path)

    # Step 2: 检测 slot（template vs variable）
    slot_result = dts.detect_slots(pptx_path, mode=mode)

    # Step 3: 整合输出
    combined = {
        "source": str(Path(pptx_path).resolve()),
        "template_meta": template_meta,  # 旧字段，保留兼容
        "global": template_meta.get("global"),
        "theme": template_meta.get("theme"),
        "layouts": template_meta.get("layouts"),
        "stats": template_meta.get("stats"),
        # 新增 slot 检测
        "slot_detection": {
            "mode": mode,
            "slides": slot_result["slides"],
            "decision": slot_result["decision"],
        },
    }

    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(combined, f, ensure_ascii=False, indent=2)

    print(f"✅ 整合分析完成：{output_path}")
    print(f"   模板版式：{len(combined['layouts'])} 个")
    print(f"   示例 slide：{len(combined['slot_detection']['slides'])} 个")
    print(f"   slot 检测模式：{mode}")


if __name__ == "__main__":
    main()
