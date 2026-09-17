#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
organic_polish.py - 内容驱动的去模板化与有机排版美化引擎
(Content-Driven Organic Presentation Evolution Engine)

核心特性：
1. 【去模板化 (De-templatization)】：摒弃千篇一律的死板方框，根据内容语义与长度自适应呼吸留白；
2. 【核心数据英雄化 (Hero Metric Amplification)】：精准识别准确率、成本倍数、损失金额等关键KPI并跳跃式放大至 30~34pt，赋予视觉锚点；
3. 【轻量化容器与微妙边界 (Airy Containers)】：消灭笨重深色边框，采用微浅色呼吸底（#F8FAFC）与 0.75pt 精细浅线（#E2E8F0）；
4. 【动态排版呼吸感 (Breathing Typography)】：正文字号全面归一到 11.5~12pt 舒适区，拉开 1.15 倍呼吸行距与段落间歇；
5. 【差异化高亮 (Differentiated Moat)】：竞品与核心结论页做非对称对比强化，让核心优势自然浮起。
"""

import sys
import os
import re
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

# 核心有机调色板 (深邃科技与自然呼吸: 科技蓝 #0066FF, 墨黑 #0F172A, 辅助蓝灰 #475569, 轻质面 #F8FAFC, 细微线 #E2E8F0, 警示/亮点红 #DC2626)
COLOR_HERO_ACCENT = RGBColor(0, 102, 255)       # 核心亮点蓝
COLOR_ALERT_ACCENT = RGBColor(220, 38, 38)       # 痛点/损失警示红
COLOR_TITLE_DARK = RGBColor(15, 23, 42)          # 大标题深黑
COLOR_BODY_SLATE = RGBColor(51, 65, 85)          # 正文蓝灰（高可读性且不生硬）
COLOR_MUTED_LABEL = RGBColor(100, 116, 139)      # 次要标注灰
COLOR_CARD_SURFACE = RGBColor(248, 250, 252)     # 轻盈呼吸卡片底
COLOR_CARD_BORDER = RGBColor(226, 232, 240)      # 0.75pt 精致浅边框
COLOR_MOAT_HIGHLIGHT = RGBColor(238, 246, 255)   # 核心优势浅蓝高亮底

METRIC_PATTERN = re.compile(r"^([±\+\-]?[0-9]+(?:\.[0-9]+)?\s*[%％]|±[0-9]+(?:\.[0-9]+)?\s*mm|[0-9]+[×xX]|1\/[0-9]+(\s*-\s*1\/[0-9]+)?|[0-9]+(?:\.[0-9]+)?\s*(?:亿元|万元|吨|个\/小时))$")

def is_hero_metric_candidate(text: str) -> bool:
    cleaned = text.strip()
    if not cleaned:
        return False
    if METRIC_PATTERN.match(cleaned):
        return True
    # 额外匹配如 "+40 %", "-20 %", "30→15 %以下", "98.6%", "±0.3mm"
    if ("%" in cleaned or "％" in cleaned or "±" in cleaned or "亿元" in cleaned) and len(cleaned) <= 15:
        # 必须含有数字
        if any(c.isdigit() for c in cleaned):
            return True
    return False

def is_alert_metric_candidate(text: str) -> bool:
    cleaned = text.strip()
    return "3000 亿" in cleaned or "3000亿" in cleaned or "28.7%" in cleaned or "损耗" in cleaned or "痛点" in cleaned

def organic_polish_presentation(
    input_path: str,
    output_path: str,
    font_family: str = "微软雅黑",
    boost_metrics: bool = True,
    soften_cards: bool = True,
    align_grid: bool = True
) -> Dict[str, Any]:
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"Input PPTX not found: {input_path}")

    prs = pptx.Presentation(input_path)
    total_slides = len(prs.slides)

    stats = {
        "slides_processed": total_slides,
        "hero_metrics_boosted": 0,
        "cards_softened": 0,
        "fonts_refined": 0,
        "tiny_fonts_fixed": 0,
        "grid_snaps": 0
    }

    for idx, slide in enumerate(prs.slides, start=1):
        text_shapes_boxes = []

        # 遍历所有形状做去模板化与有机升级
        for shape in slide.shapes:
            # 1. 软化卡片容器 (去模板化：消除生硬的深灰厚重块，代以 0.75pt 轻柔浅线与透气呼吸底)
            if soften_cards and shape.shape_type == MSO_SHAPE_TYPE.AUTO_SHAPE:
                shape_name_lower = shape.name.lower()
                # 针对背景卡片/容器形状
                if ("rounded rectangle" in shape_name_lower or "rectangle" in shape_name_lower) and shape.width.pt > 60 and shape.height.pt > 20:
                    # 排除全屏背景矩形 (w>=950 and h>=530) 以及极细装饰线 (h<=6 or w<=6)
                    if not (shape.width.pt >= 950 and shape.height.pt >= 530) and shape.height.pt > 8 and shape.width.pt > 8:
                        try:
                            # 调整边框为精细线
                            shape.line.width = Pt(0.75)
                            shape.line.color.rgb = COLOR_CARD_BORDER
                            # 针对竞品对标页 (Slide 11) 的"本产线"列给予专属微蓝高光
                            if idx == 11 and 240 <= shape.left.pt <= 405 and shape.width.pt <= 165:
                                shape.fill.solid()
                                shape.fill.fore_color.rgb = COLOR_MOAT_HIGHLIGHT
                                shape.line.color.rgb = COLOR_HERO_ACCENT
                            stats["cards_softened"] += 1
                        except Exception:
                            pass

            # 2. 文本框排版与核心数据英雄化处理
            if shape.has_text_frame:
                text_shapes_boxes.append((shape.left.pt, shape.top.pt, shape.width.pt, shape.height.pt, shape))
                tf = shape.text_frame
                tf.word_wrap = True

                # 优化内边距释放空间
                tf.margin_top = Pt(4)
                tf.margin_bottom = Pt(4)
                tf.margin_left = Pt(6)
                tf.margin_right = Pt(6)

                raw_full_text = "".join(p.text for p in tf.paragraphs).strip()
                is_hero = is_hero_metric_candidate(raw_full_text)

                for p in tf.paragraphs:
                    p_text = p.text.strip()
                    # 呼吸行间距微调
                    p.line_spacing = 1.15
                    p.space_after = Pt(3)

                    for r in p.runs:
                        r_text = r.text.strip()
                        # 统一字体族
                        r.font.name = font_family
                        stats["fonts_refined"] += 1

                        # 核心指标英雄化放大
                        if boost_metrics and (is_hero or is_hero_metric_candidate(r_text)):
                            r.font.size = Pt(30.0)
                            r.font.bold = True
                            if is_alert_metric_candidate(raw_full_text):
                                r.font.color.rgb = COLOR_ALERT_ACCENT
                            else:
                                r.font.color.rgb = COLOR_HERO_ACCENT
                            stats["hero_metrics_boosted"] += 1
                        else:
                            # 字号层级阶梯与过小字号治理
                            if r.font.size:
                                f_size = r.font.size.pt
                                if f_size < 11.5:
                                    r.font.size = Pt(11.5)
                                    stats["tiny_fonts_fixed"] += 1
                                # 封面大标题呼吸空间强化 (Slide 1)
                                if idx == 1 and f_size >= 28:
                                    r.font.size = Pt(44.0)
                                    r.font.bold = True
                                    r.font.color.rgb = COLOR_TITLE_DARK

            # 3. 表格内文本优化 (如果有原生表格)
            elif shape.has_table:
                for row in shape.table.rows:
                    for cell in row.cells:
                        cell.text_frame.word_wrap = True
                        cell.text_frame.margin_top = Pt(3)
                        cell.text_frame.margin_bottom = Pt(3)
                        for p in cell.text_frame.paragraphs:
                            for r in p.runs:
                                r.font.name = font_family
                                if r.font.size and r.font.size.pt < 11.0:
                                    r.font.size = Pt(11.5)
                                    stats["tiny_fonts_fixed"] += 1

        # 4. 智能网格对齐吸附 (消除 1~12pt 的并排微抖动)
        if align_grid and len(text_shapes_boxes) >= 2:
            sorted_boxes = sorted(text_shapes_boxes, key=lambda b: b[1])
            groups = []
            curr = [sorted_boxes[0]]
            for i in range(1, len(sorted_boxes)):
                item = sorted_boxes[i]
                if abs(item[1] - curr[0][1]) <= 12.0:
                    curr.append(item)
                else:
                    if len(curr) > 1:
                        groups.append(curr)
                    curr = [item]
            if len(curr) > 1:
                groups.append(curr)

            for g in groups:
                target_top = min(item[1] for item in g)
                for item in g:
                    shape_obj = item[4]
                    if abs(shape_obj.top.pt - target_top) > 0.5:
                        shape_obj.top = Pt(target_top)
                        stats["grid_snaps"] += 1

    prs.save(output_path)
    return stats

def main():
    parser = argparse.ArgumentParser(description="Organic Content-Driven Presentation Polisher")
    parser.add_argument("input_pptx", help="Path to input PPTX")
    parser.add_argument("-o", "--output", required=True, help="Path to output polished PPTX")
    parser.add_argument("--font", default="微软雅黑", help="Primary font family (default: 微软雅黑)")
    args = parser.parse_args()

    print("=" * 70)
    print(f"🌱 正在启动【内容驱动·去模板化】有机排版进化引擎...")
    print(f"   输入: {os.path.basename(args.input_path if hasattr(args, 'input_path') else args.input_pptx)}")
    print(f"   输出: {os.path.basename(args.output)}")
    print("=" * 70)

    stats = organic_polish_presentation(
        input_path=args.input_pptx,
        output_path=args.output,
        font_family=args.font,
        boost_metrics=True,
        soften_cards=True,
        align_grid=True
    )

    print("\n✅ 有机排版进化完成！核心指标:")
    print(f"  - 处理幻灯片: {stats['slides_processed']} 页")
    print(f"  - 核心数据英雄化放大 (Hero Metrics): {stats['hero_metrics_boosted']} 处")
    print(f"  - 容器去模板化与呼吸软化: {stats['cards_softened']} 个容器")
    print(f"  - 字体规范与呼吸行距: {stats['fonts_refined']} 处")
    print(f"  - 极小字号梯级治理 (<11.5pt): {stats['tiny_fonts_fixed']} 处")
    print(f"  - 水平网格微抖动平滑吸附: {stats['grid_snaps']} 处")
    print(f"\n📁 进阶交付产物已生成: {args.output}")
    print("=" * 70)

if __name__ == "__main__":
    main()
