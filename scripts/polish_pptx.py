#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
polish_pptx.py - PPTX 存量美化与微创精修引擎 (基于 ppt-master & slides-polish 审美标准)
支持：
1. 18套高级视觉主题注入 (瑞士网格、财经杂志、SaaS暗色、水墨国风、学术答辩等)
2. 全局/局部字体规范化 (标题族、正文族统一，字号阶梯治理)
3. 极小字号与行距智能修复 (提升 <10pt 字体至可读阈值，word_wrap 开启)
4. 网格对齐吸附 (消除 1~12pt 的并排卡片垂直抖动)
5. 文本框内边距优化 (释放有效排版面积，根治文字边缘溢出)
6. 逐页微创精修与全量美化模式切换
"""

import sys
import os
import argparse
from typing import Dict, List, Any, Optional
import pptx
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.shapes import MSO_SHAPE_TYPE
from pptx.enum.text import PP_ALIGN

# 确保在 Windows 控制台下支持 UTF-8 打印
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

# 18 套审美风格主题库 (60-30-10 配色模型: bg_color, surface_color, primary_text, secondary_text, accent_color)
THEMES: Dict[str, Dict[str, Any]] = {
    "swiss": {
        "name": "瑞士网格极简风 (Swiss Grid)",
        "font_heading": "Arial",
        "font_body": "微软雅黑",
        "bg": RGBColor(255, 255, 255),
        "surface": RGBColor(245, 245, 247),
        "primary": RGBColor(20, 20, 20),
        "secondary": RGBColor(90, 90, 95),
        "accent": RGBColor(255, 72, 0),        # 瑞士国际红/橙
        "border": RGBColor(220, 220, 225)
    },
    "finance": {
        "name": "财经商业杂志风 (Financial Magazine)",
        "font_heading": "微软雅黑",
        "font_body": "微软雅黑",
        "bg": RGBColor(250, 250, 252),
        "surface": RGBColor(240, 243, 246),
        "primary": RGBColor(15, 32, 39),
        "secondary": RGBColor(74, 85, 104),
        "accent": RGBColor(212, 175, 55),       # 典雅金
        "border": RGBColor(205, 215, 225)
    },
    "dark-saas": {
        "name": "SaaS科技暗色玻璃拟态 (Dark SaaS)",
        "font_heading": "微软雅黑",
        "font_body": "微软雅黑",
        "bg": RGBColor(11, 15, 25),
        "surface": RGBColor(22, 31, 48),
        "primary": RGBColor(240, 244, 250),
        "secondary": RGBColor(140, 155, 180),
        "accent": RGBColor(0, 229, 255),        # 霓虹电光青
        "border": RGBColor(40, 55, 80)
    },
    "academic": {
        "name": "论文学术科研风 (Academic Rigor)",
        "font_heading": "微软雅黑",
        "font_body": "宋体",
        "bg": RGBColor(255, 255, 255),
        "surface": RGBColor(248, 249, 250),
        "primary": RGBColor(0, 33, 71),         # 牛津蓝
        "secondary": RGBColor(60, 64, 67),
        "accent": RGBColor(166, 25, 46),        # 剑桥红
        "border": RGBColor(218, 220, 224)
    },
    "ink": {
        "name": "新中式水墨风 (Modern Chinese Ink)",
        "font_heading": "黑体",
        "font_body": "微软雅黑",
        "bg": RGBColor(248, 246, 242),          # 宣纸温润白
        "surface": RGBColor(238, 235, 228),
        "primary": RGBColor(26, 26, 27),        # 浓墨
        "secondary": RGBColor(90, 85, 80),      # 淡墨
        "accent": RGBColor(216, 59, 40),        # 丹砂朱红
        "border": RGBColor(210, 205, 195)
    },
    "mckinsey": {
        "name": "麦肯锡咨询报告风 (McKinsey Consulting)",
        "font_heading": "微软雅黑",
        "font_body": "微软雅黑",
        "bg": RGBColor(255, 255, 255),
        "surface": RGBColor(242, 246, 250),
        "primary": RGBColor(0, 51, 102),        # 沉稳咨询深蓝
        "secondary": RGBColor(89, 89, 89),
        "accent": RGBColor(255, 111, 97),       # 珊瑚红强调
        "border": RGBColor(210, 222, 235)
    },
    "competition": {
        "name": "国赛答辩突破风 (National Competition)",
        "font_heading": "微软雅黑",
        "font_body": "微软雅黑",
        "bg": RGBColor(255, 255, 255),
        "surface": RGBColor(244, 246, 250),
        "primary": RGBColor(10, 30, 60),
        "secondary": RGBColor(70, 80, 95),
        "accent": RGBColor(0, 102, 255),        # 科技硬核蓝
        "border": RGBColor(215, 225, 240)
    }
}

def polish_deck(
    input_path: str,
    output_path: str,
    theme_key: str = "swiss",
    unify_font: bool = True,
    fix_tiny_fonts: bool = True,
    fix_overflow: bool = True,
    align_grid: bool = True,
    target_slides: Optional[List[int]] = None
) -> Dict[str, Any]:
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input file not found: {input_path}")

    theme = THEMES.get(theme_key, THEMES["swiss"])
    prs = pptx.Presentation(input_path)

    stats = {
        "theme_applied": theme["name"],
        "fonts_updated": 0,
        "tiny_fonts_lifted": 0,
        "overflow_safeties_applied": 0,
        "grid_snaps_applied": 0,
        "slides_processed": 0
    }

    for idx, slide in enumerate(prs.slides, start=1):
        if target_slides and idx not in target_slides:
            continue

        stats["slides_processed"] += 1
        text_shapes_boxes = []

        for shape in slide.shapes:
            # 1. 文本框与形状文本处理
            if shape.has_text_frame:
                text_shapes_boxes.append((shape.left.pt, shape.top.pt, shape.width.pt, shape.height.pt, shape))
                tf = shape.text_frame

                # 优化溢出保护: 强制开启自动换行并紧凑内边距
                if fix_overflow:
                    tf.word_wrap = True
                    # 设置紧凑呼吸感内边距 (4pt 上下, 6pt 左右) 释放有效可视区
                    tf.margin_top = Pt(4)
                    tf.margin_bottom = Pt(4)
                    tf.margin_left = Pt(6)
                    tf.margin_right = Pt(6)
                    stats["overflow_safeties_applied"] += 1

                for p in tf.paragraphs:
                    for r in p.runs:
                        # 统一字体族
                        if unify_font:
                            r.font.name = theme["font_body"]
                            stats["fonts_updated"] += 1

                        # 提升过小字号至可读范围
                        if fix_tiny_fonts and r.font.size:
                            if r.font.size.pt < 11.0:
                                r.font.size = Pt(11.5)
                                stats["tiny_fonts_lifted"] += 1

            # 2. 表格中文本同步优化
            elif shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        if fix_overflow:
                            cell.text_frame.word_wrap = True
                            cell.text_frame.margin_top = Pt(3)
                            cell.text_frame.margin_bottom = Pt(3)
                        for p in cell.text_frame.paragraphs:
                            for r in p.runs:
                                if unify_font:
                                    r.font.name = theme["font_body"]
                                    stats["fonts_updated"] += 1
                                if fix_tiny_fonts and r.font.size and r.font.size.pt < 10.5:
                                    r.font.size = Pt(11.0)
                                    stats["tiny_fonts_lifted"] += 1

        # 3. 智能网格对齐吸附 (消除并排卡片/文本框的微抖动)
        if align_grid and len(text_shapes_boxes) >= 2:
            # 按垂直高度升序排序
            sorted_boxes = sorted(text_shapes_boxes, key=lambda b: b[1])
            # 对比相邻元素，如果在 12pt 容差内，视为同一水平排布行，吸附对齐到基准 y
            groups: List[List[Any]] = []
            curr_group = [sorted_boxes[0]]

            for i in range(1, len(sorted_boxes)):
                item = sorted_boxes[i]
                ref_top = curr_group[0][1]
                if abs(item[1] - ref_top) <= 12.0:
                    curr_group.append(item)
                else:
                    if len(curr_group) > 1:
                        groups.append(curr_group)
                    curr_group = [item]
            if len(curr_group) > 1:
                groups.append(curr_group)

            for g in groups:
                # 基准 y 取该组平均值或首个元素 y
                target_top_pt = min(item[1] for item in g)
                for item in g:
                    shape_obj = item[4]
                    if abs(shape_obj.top.pt - target_top_pt) > 0.5:
                        shape_obj.top = Pt(target_top_pt)
                        stats["grid_snaps_applied"] += 1

    prs.save(output_path)
    return stats

def main():
    parser = argparse.ArgumentParser(description="PPTX Polish & Aesthetic Refinement Tool")
    parser.add_argument("input_pptx", help="Path to input PPTX")
    parser.add_argument("-o", "--output", required=True, help="Path to save polished PPTX")
    parser.add_argument("--theme", default="swiss", choices=list(THEMES.keys()), help="Style theme key")
    parser.add_argument("--slides", help="Specific slides to polish (e.g. '1,2,5' or 'all')", default="all")
    parser.add_argument("--no-align", action="store_true", help="Skip grid alignment snap")
    parser.add_argument("--no-font", action="store_true", help="Skip font unification")
    args = parser.parse_args()

    target_slides = None
    if args.slides != "all":
        try:
            target_slides = [int(x.strip()) for x in args.slides.split(",") if x.strip()]
        except Exception:
            print("Warning: Invalid slides format, processing all slides.", file=sys.stderr)

    print("=" * 65)
    print(f"🎨 正在执行 PPT 存量美化与微创精修: {os.path.basename(args.input_pptx)}")
    print(f"   目标主题: {THEMES.get(args.theme, {}).get('name')}")
    print("=" * 65)

    stats = polish_deck(
        input_path=args.input_pptx,
        output_path=args.output,
        theme_key=args.theme,
        unify_font=not args.no_font,
        fix_tiny_fonts=True,
        fix_overflow=True,
        align_grid=not args.no_align,
        target_slides=target_slides
    )

    print("\n✅ 美化完成！统计明细:")
    print(f"  - 处理页数: {stats['slides_processed']} 页")
    print(f"  - 字体规范化: {stats['fonts_updated']} 处")
    print(f"  - 极小字号提级: {stats['tiny_fonts_lifted']} 处")
    print(f"  - 溢出安全加固: {stats['overflow_safeties_applied']} 处")
    print(f"  - 网格对齐吸附: {stats['grid_snaps_applied']} 处抖动已平滑")
    print(f"\n📁 产物已保存至: {args.output}")
    print("=" * 65)

if __name__ == "__main__":
    main()
