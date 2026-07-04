#!/usr/bin/env python3
"""
extract_layouts.py
基于 parse_template.py 的输出，做高层语义识别：
- 为 AI 提供"哪一版式适合做哪类页面"的友好描述
- 提取装饰元素特征（logo 位置、背景、配色）
- 输出页面类型选择指南 + 装饰元素清单

Usage:
    python extract_layouts.py <template_meta.json> <output.json>
"""
import json
import sys
from collections import defaultdict
from pathlib import Path


# 页面类型与版式分类的对应关系
PAGE_TYPE_HINTS = {
    "cover": {
        "cn": "封面页",
        "purpose": "全文档首页，展示标题与副标题",
        "layout_types": ["cover"],
        "fallback_names": ["Title Slide", "封面", "Cover", "标题页"],
    },
    "toc": {
        "cn": "目录页",
        "purpose": "展示内容大纲与分章节",
        "layout_types": ["toc"],
        "fallback_names": ["Table of Contents", "目录", "Agenda", "TOC"],
    },
    "section_header": {
        "cn": "章节首页",
        "purpose": "分章节过渡，承上启下",
        "layout_types": ["section_header"],
        "fallback_names": ["Section Header", "章节", "Section"],
    },
    "content": {
        "cn": "正文页",
        "purpose": "核心内容展示，标题+要点+图表/表格",
        "layout_types": ["content"],
        "fallback_names": ["Title and Content", "Content", "正文", "标题与内容"],
    },
    "summary": {
        "cn": "总结页",
        "purpose": "关键结论/致谢/封底",
        "layout_types": ["summary"],
        "fallback_names": ["Thank You", "谢谢", "Summary", "Conclusion", "Q&A"],
    },
}

# 兜底策略：哪些页面类型可以回退到"占位符最多"的版式
FALLBACK_TO_RICHEST = {"content", "toc", "summary"}
# 哪些类型可以回退到"占位符最少（纯标题）"的版式
FALLBACK_TO_SECTION = {"section_header"}


def find_layout_for_type(layouts, type_name, page_type_hint):
    """根据 hint 在 layouts 中找最合适的版式"""
    hint = page_type_hint
    # 1) 优先按已分类的 type 匹配
    for layout in layouts:
        if layout["type"] == type_name:
            return layout
    # 2) 按名称模糊匹配
    for layout in layouts:
        if any(
            kw.lower() in (layout["name"] or "").lower()
            for kw in hint["fallback_names"]
        ):
            return layout
    # 3) 兜底策略
    if type_name in FALLBACK_TO_RICHEST:
        max_ph = -1
        best = None
        for layout in layouts:
            # 避免选到 cover 类型
            if layout["type"] == "cover":
                continue
            n = len(layout["placeholders"])
            if n > max_ph:
                max_ph = n
                best = layout
        if best:
            return best
    if type_name in FALLBACK_TO_SECTION:
        # 找占位符少的（一般是 section header）
        min_ph = 999
        best = None
        for layout in layouts:
            if layout["type"] == "cover":
                continue
            n = len(layout["placeholders"])
            if 0 < n < min_ph:
                min_ph = n
                best = layout
        if best:
            return best
    if type_name == "cover":
        return layouts[0] if layouts else None
    return None


def describe_position(ph):
    """人类可读的位置描述"""
    pos = ph.get("position", {})
    left = pos.get("left_pt")
    top = pos.get("top_pt")
    w = pos.get("width_pt")
    h = pos.get("height_pt")
    if left is None or top is None:
        return "位置未知"
    # 简单分区
    vertical = "顶部" if top < 200 else ("底部" if top > 500 else "中部")
    horizontal = "左侧" if left < 300 else ("右侧" if left > 700 else "中间")
    return f"{vertical}{horizontal}（约 {w:.0f}×{h:.0f} pt）"


def describe_layout_for_ai(layout):
    """为 AI 输出友好版式描述"""
    if layout is None:
        return None
    desc = {
        "name": layout["name"],
        "type": layout["type"],
        "placeholders": [],
        "decorative_count": len(layout.get("decorative_shapes", [])),
    }
    for ph in layout["placeholders"]:
        ph_info = {
            "type": ph["type"],
            "name": ph["name"],
            "position": describe_position(ph),
            "sample_text": ph.get("sample_text"),
            "font": ph.get("font"),
        }
        desc["placeholders"].append(ph_info)
    return desc


def build_layout_guide(template_meta):
    """构建版式选择指南"""
    layouts = template_meta["layouts"]
    guide = {}
    for type_name, hint in PAGE_TYPE_HINTS.items():
        chosen = find_layout_for_type(layouts, type_name, hint)
        guide[type_name] = {
            "cn_name": hint["cn"],
            "purpose": hint["purpose"],
            "recommended_layout": chosen["name"] if chosen else None,
            "description": describe_layout_for_ai(chosen) if chosen else None,
        }

    # 装饰元素清单
    decorations = []
    for layout in layouts:
        for shape in layout.get("decorative_shapes", []):
            if shape.get("is_picture"):
                decorations.append(
                    {
                        "layout": layout["name"],
                        "type": "logo/picture",
                        "position": shape["position"],
                    }
                )
            elif shape.get("fill_color"):
                decorations.append(
                    {
                        "layout": layout["name"],
                        "type": "color_shape",
                        "color": shape["fill_color"],
                        "position": shape["position"],
                    }
                )
    return {"layout_guide": guide, "decorations": decorations}


def main():
    if len(sys.argv) != 3:
        print("Usage: python extract_layouts.py <template_meta.json> <output.json>")
        sys.exit(1)
    input_path = sys.argv[1]
    output_path = sys.argv[2]
    with open(input_path, "r", encoding="utf-8") as f:
        template_meta = json.load(f)
    result = build_layout_guide(template_meta)
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(result, f, ensure_ascii=False, indent=2)
    print(f"✅ 版式指南生成完成：{output_path}")
    for t, info in result["layout_guide"].items():
        print(f"   {info['cn_name']:8s} -> {info['recommended_layout']}")


if __name__ == "__main__":
    main()
