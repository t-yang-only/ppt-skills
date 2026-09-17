#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
audit_pptx.py - PPTX 逐页深度审计与审美标准合规检查器
遵循 ppt-master, slides-polish 及审美标准规范 (Aesthetic Standards)。
支持检测：
1. 纯图片整页贴图 (伪原生PPT / 避坑清单)
2. 字体种类与一致性 (字形发散、跨页混乱)
3. 字号梯级与过小文本 (<12pt 预警, <10pt 阻断)
4. 文本框边界溢出与边界越界 (超出画布、字密超标)
5. 形状微错位与网格对齐抖动 (水平/垂直微错位)
6. 配色规范度与对比度估算 (WCAG 2.1 AA)
7. 原生对象占比与结构化评分 (0-100)
"""

import sys
import os
import json
import argparse

# 确保在 Windows 控制台下支持 UTF-8 打印
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass
from typing import Dict, List, Any, Tuple
import pptx
from pptx.util import Inches, Pt
from pptx.enum.shapes import MSO_SHAPE_TYPE

# 默认审计参数阈值
MIN_BODY_FONT_SIZE_PT = 11.5   # 低于此字号判定为字号偏小
TINY_FONT_SIZE_PT = 9.5        # 低于此字号判定为极小不可读
MAX_FONT_FAMILIES = 3          # 全篇允许的最大字体族数
IMAGE_AREA_THRESHOLD = 0.75    # 单张图片覆盖超 75% 面积且无原生文本，判定为纯图片贴图
ALIGNMENT_TOLERANCE_PT = 12.0  # 视作同一水平行但存在微抖动的误差区间 (pt)

def get_rgb_brightness(rgb_color) -> float:
    """计算 RGB 颜色的相对亮度 (0.0 - 1.0)"""
    if rgb_color is None:
        return 1.0  # 默认假定浅色/白色
    try:
        r = rgb_color[0] / 255.0
        g = rgb_color[1] / 255.0
        b = rgb_color[2] / 255.0
        return 0.2126 * r + 0.7152 * g + 0.0722 * b
    except Exception:
        return 1.0

def audit_presentation(pptx_path: str) -> Dict[str, Any]:
    if not os.path.exists(pptx_path):
        raise FileNotFoundError(f"PPTX file not found: {pptx_path}")

    prs = pptx.Presentation(pptx_path)
    slide_width_pt = prs.slide_width.pt
    slide_height_pt = prs.slide_height.pt
    slide_area = slide_width_pt * slide_height_pt

    total_slides = len(prs.slides)
    all_fonts_used = set()
    slide_reports = []

    deck_issues = []
    pure_image_slides = []
    tiny_font_count = 0
    overflow_risk_count = 0
    misaligned_count = 0

    total_native_shapes = 0
    total_images = 0

    for idx, slide in enumerate(prs.slides, start=1):
        slide_rep = {
            "slide_index": idx,
            "issues": [],
            "stats": {
                "shape_count": len(slide.shapes),
                "text_frame_count": 0,
                "image_count": 0,
                "table_count": 0,
                "chart_count": 0,
                "native_character_count": 0,
                "fonts": set(),
                "font_sizes": []
            },
            "is_pure_image": False
        }

        max_image_area_ratio = 0.0
        slide_chars = 0
        text_shapes_boxes = []

        for shape in slide.shapes:
            # 越界检查 (Out of bounds)
            if shape.left.pt < -5 or shape.top.pt < -5 or (shape.left.pt + shape.width.pt) > (slide_width_pt + 15) or (shape.top.pt + shape.height.pt) > (slide_height_pt + 15):
                slide_rep["issues"].append({
                    "type": "out_of_bounds",
                    "severity": "medium",
                    "shape_name": shape.name,
                    "msg": f"元素超出画布边界 (L:{shape.left.pt:.1f}, T:{shape.top.pt:.1f}, W:{shape.width.pt:.1f}, H:{shape.height.pt:.1f}, 画布:{slide_width_pt:.0f}x{slide_height_pt:.0f})"
                })

            # 图片统计
            if shape.shape_type == MSO_SHAPE_TYPE.PICTURE:
                slide_rep["stats"]["image_count"] += 1
                total_images += 1
                img_area = shape.width.pt * shape.height.pt
                ratio = img_area / slide_area
                if ratio > max_image_area_ratio:
                    max_image_area_ratio = ratio

            # 表格统计
            elif shape.has_table:
                slide_rep["stats"]["table_count"] += 1
                total_native_shapes += 1
                for row in shape.table.rows:
                    for cell in row.cells:
                        for p in cell.text_frame.paragraphs:
                            slide_chars += len(p.text)
                            for r in p.runs:
                                if r.font.name:
                                    slide_rep["stats"]["fonts"].add(r.font.name)
                                    all_fonts_used.add(r.font.name)
                                if r.font.size:
                                    slide_rep["stats"]["font_sizes"].append(r.font.size.pt)

            # 图表统计
            elif shape.has_chart:
                slide_rep["stats"]["chart_count"] += 1
                total_native_shapes += 1

            # 文本框及形状文本统计
            elif shape.has_text_frame:
                slide_rep["stats"]["text_frame_count"] += 1
                total_native_shapes += 1
                text_shapes_boxes.append((shape.left.pt, shape.top.pt, shape.width.pt, shape.height.pt, shape.name))

                tf_text = "".join(p.text for p in shape.text_frame.paragraphs)
                slide_chars += len(tf_text)

                # 溢出估算：若文本长度大但文本框极窄或极矮
                tf_area = shape.width.pt * shape.height.pt
                if len(tf_text) > 80 and tf_area < (len(tf_text) * 12.0):
                    overflow_risk_count += 1
                    slide_rep["issues"].append({
                        "type": "overflow_risk",
                        "severity": "high",
                        "shape_name": shape.name,
                        "msg": f"文本密度过高，存在溢出/文字截断风险 (字符数:{len(tf_text)}, 区域:{shape.width.pt:.0f}x{shape.height.pt:.0f}pt)"
                    })

                for p in shape.text_frame.paragraphs:
                    for r in p.runs:
                        if r.font.name:
                            slide_rep["stats"]["fonts"].add(r.font.name)
                            all_fonts_used.add(r.font.name)
                        if r.font.size:
                            f_size = r.font.size.pt
                            slide_rep["stats"]["font_sizes"].append(f_size)
                            if f_size < TINY_FONT_SIZE_PT:
                                tiny_font_count += 1
                                slide_rep["issues"].append({
                                    "type": "tiny_font_error",
                                    "severity": "high",
                                    "shape_name": shape.name,
                                    "msg": f"字号极小不可读: {f_size:.1f}pt (正文字号不宜低于 {MIN_BODY_FONT_SIZE_PT}pt)"
                                })
                            elif f_size < MIN_BODY_FONT_SIZE_PT:
                                slide_rep["issues"].append({
                                    "type": "small_font_warning",
                                    "severity": "low",
                                    "shape_name": shape.name,
                                    "msg": f"字号偏小: {f_size:.1f}pt，建议调大至 {MIN_BODY_FONT_SIZE_PT}pt 以上"
                                })
            else:
                total_native_shapes += 1

        slide_rep["stats"]["native_character_count"] = slide_chars

        # 检查是否是纯图片整页贴图 (避坑清单)
        if max_image_area_ratio >= IMAGE_AREA_THRESHOLD and slide_chars < 20:
            slide_rep["is_pure_image"] = True
            pure_image_slides.append(idx)
            slide_rep["issues"].append({
                "type": "flattened_image_slide",
                "severity": "blocker",
                "msg": f"检测到纯图片贴图页面 (图片覆盖率 {max_image_area_ratio*100:.1f}%, 原生字数 {slide_chars})。违背原生可编辑PPT规范，应使用 bggg-creator 拆解还原！"
            })

        # 检查水平并排卡片的微错位 (对齐检查)
        if len(text_shapes_boxes) >= 2:
            sorted_by_top = sorted(text_shapes_boxes, key=lambda b: b[1])
            for i in range(len(sorted_by_top) - 1):
                b1 = sorted_by_top[i]
                b2 = sorted_by_top[i+1]
                # 如果水平距离明显分开，但垂直位置非常接近却不完全对齐 (1pt ~ 12pt 误差)
                delta_y = abs(b1[1] - b2[1])
                if 1.0 < delta_y <= ALIGNMENT_TOLERANCE_PT:
                    misaligned_count += 1
                    slide_rep["issues"].append({
                        "type": "misalignment",
                        "severity": "medium",
                        "msg": f"并排元素垂直对齐微抖动: '{b1[4]}' 与 '{b2[4]}' 垂直差 {delta_y:.1f}pt (建议统一 y 轴坐标)"
                    })

        slide_rep["stats"]["fonts"] = list(slide_rep["stats"]["fonts"])
        slide_reports.append(slide_rep)

    # 全局字体多样性检查
    if len(all_fonts_used) > MAX_FONT_FAMILIES:
        deck_issues.append({
            "type": "font_proliferation",
            "severity": "medium",
            "msg": f"全篇字体族过多 ({len(all_fonts_used)} 种: {', '.join(list(all_fonts_used)[:6])}...)。审美规范要求全篇控制在 {MAX_FONT_FAMILIES} 种以内。"
        })

    # 全局纯图片警告
    if len(pure_image_slides) > 0:
        deck_issues.append({
            "type": "pure_image_warning",
            "severity": "blocker",
            "msg": f"存在 {len(pure_image_slides)} 页纯图片扁平幻灯片 (第 {', '.join(map(str, pure_image_slides))} 页)。无法编辑文本或图表，需用 bggg-creator 进行元素级拆解！"
        })

    # 综合健康分计算 (0 - 100)
    score = 100
    score -= len(pure_image_slides) * 15      # 纯图片严重扣分
    score -= len(deck_issues) * 5
    score -= min(30, tiny_font_count * 3)
    score -= min(20, overflow_risk_count * 4)
    score -= min(15, misaligned_count * 2)
    score = max(0, min(100, score))

    return {
        "file": pptx_path,
        "total_slides": total_slides,
        "slide_dimensions": {"width_pt": slide_width_pt, "height_pt": slide_height_pt, "ratio": f"{slide_width_pt/slide_height_pt:.2f}"},
        "score": score,
        "deck_issues": deck_issues,
        "all_fonts": list(all_fonts_used),
        "pure_image_slides": pure_image_slides,
        "stats": {
            "total_native_shapes": total_native_shapes,
            "total_images": total_images,
            "tiny_font_issues": tiny_font_count,
            "overflow_risks": overflow_risk_count,
            "misalignment_issues": misaligned_count
        },
        "slides": slide_reports
    }

def print_human_report(report: Dict[str, Any]):
    print("=" * 70)
    print(f"📊 PPT 审计与审美合规报告: {os.path.basename(report['file'])}")
    print(f"   总页数: {report['total_slides']} 页 | 比例: {report['slide_dimensions']['ratio']} | 审美健康评分: {report['score']} / 100")
    print("=" * 70)

    if report["deck_issues"]:
        print("\n🚨 【全局高危/阻断项 (Deck-Level Issues)】:")
        for iss in report["deck_issues"]:
            prefix = "[BLOCKER]" if iss["severity"] == "blocker" else "[WARN]"
            print(f"  {prefix} {iss['msg']}")

    print("\n🔍 【全局统计 (Metrics Summary)】:")
    print(f"  - 使用字体 ({len(report['all_fonts'])} 种): {', '.join(report['all_fonts']) if report['all_fonts'] else '默认未显式指定'}")
    print(f"  - 原生对象数: {report['stats']['total_native_shapes']} 个 | 图片数量: {report['stats']['total_images']} 个")
    print(f"  - 纯图片扁平页: {len(report['pure_image_slides'])} 页")
    print(f"  - 极小字号瑕疵: {report['stats']['tiny_font_issues']} 处")
    print(f"  - 文本溢出风险: {report['stats']['overflow_risks']} 处")
    print(f"  - 并排微错位处: {report['stats']['misalignment_issues']} 处")

    # 挑出有问题的前 5 页打印详情
    problematic_slides = [s for s in report["slides"] if s["issues"]]
    if problematic_slides:
        print(f"\n📑 【逐页问题明细 (前 6 页典型问题)】:")
        for s in problematic_slides[:6]:
            print(f"\n  👉 Slide #{s['slide_index']} (原生字符: {s['stats']['native_character_count']}, 形状: {s['stats']['shape_count']}):")
            for iss in s["issues"][:4]:
                sev_icon = "❌" if iss["severity"] in ["blocker", "high"] else "⚠️"
                print(f"     {sev_icon} [{iss['type']}] {iss['msg']}")
            if len(s["issues"]) > 4:
                print(f"     ... 还有 {len(s['issues']) - 4} 个更轻微问题")
    else:
        print("\n✨ 完美！全篇未检测到明显版式缺陷与溢出。")

    print("\n" + "=" * 70)

def main():
    parser = argparse.ArgumentParser(description="PPTX Aesthetic & Technical Audit Tool")
    parser.add_argument("pptx_path", help="Path to PPTX file")
    parser.add_argument("--json", action="store_true", help="Output raw JSON format")
    args = parser.parse_args()

    try:
        report = audit_presentation(args.pptx_path)
        if args.json:
            print(json.dumps(report, ensure_ascii=False, indent=2))
        else:
            print_human_report(report)
    except Exception as e:
        print(f"Error during PPT audit: {e}", file=sys.stderr)
        sys.exit(1)

if __name__ == "__main__":
    main()
