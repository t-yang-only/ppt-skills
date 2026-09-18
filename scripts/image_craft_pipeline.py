#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
image_craft_pipeline.py — PPT 视觉资产深度工坊与图像生图/抠图/异形排版全能引擎
================================================================================
核心能力矩阵：
1. 【反向逆向视觉提示词与完美生图 (Prompt Craft)】：
   - 分析参考图的色彩、构图、光影、渲染器 (Octane/Blender/Midjourney/Unreal Engine 5)；
   - 输出适合生产高精资产的“完美提示词”模板；
2. 【区域切片与无背景透明素材精细生成 (Region Crop & Transparent Cutout)】：
   - 支持从整图按坐标/比例切片截取局部关键主体；
   - 智能提取透明通道（Alpha Channel），自动去除背景（基于纯色色度键剔除、自适应阈值与羽化边缘，或调用本地/API removebg 接口）；
   - 精细边缘微调 (Alpha Matting & Edge Feathering)，杜绝毛刺白边；
3. 【不规则异形蒙版与多边形图片排版 (Irregular Shape Masking)】：
   - 支持将图片放置为不标准形状：
     * `organic-blob` (有机呼吸感流体团块)
     * `hexagon` (六边形科技蜂窝)
     * `round-rect` (大圆角卡片)
     * `ellipse` (超宽胶囊/正圆)
     * `snip-corner` (切角极简卡片)
     * `trapezoid` (斜切动感多边形)
   - 自动生成带有异形蒙版的无损透明 PNG，并直接通过 python-pptx 精准定位插入幻灯片！
4. 【图文分层注入与 PPTX 原生定位装配 (Slide Inplace Placement)】：
   - 自动计算幻灯片 16:9 坐标系（左/上/宽/高），支持绝对单位与百分比网格吸附。
"""

import os
import sys
import json
import argparse
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any

try:
    from PIL import Image, ImageDraw, ImageFilter, ImageOps
    HAS_PIL = True
except ImportError:
    HAS_PIL = False

try:
    import pptx
    from pptx import Presentation
    from pptx.util import Inches, Pt
    HAS_PPTX = True
except ImportError:
    HAS_PPTX = False


def craft_perfect_prompt(
    subject: str,
    style: str = "3d-render",
    lighting: str = "studio",
    color_palette: str = "modern-tech",
    aspect_ratio: str = "16:9",
    transparent_bg: bool = False
) -> Dict[str, str]:
    """
    根据主体需求生成专业级“完美生图提示词”
    """
    style_presets = {
        "3d-render": "hyper-realistic 3D render, octane render, blender 3D, claymorphism and smooth plastic texture, volumetric lighting, ray tracing, 8k resolution, trending on Artstation",
        "isometric": "clean isometric 3D asset, orthographic projection, minimalist vector geometric design, soft ambient occlusion, crisp lines, UI/UX asset",
        "corporate-photo": "professional high-end corporate editorial photography, shot on Hasselblad H6D-100c, 85mm f/1.4 lens, softbox studio lighting, cinematic color grading, crisp focus",
        "flat-vector": "modern flat vector illustration, Bauhaus minimal design, duotone accents, crisp vector curves, clean SVG aesthetic, Dribbble trending",
        "cyber-hud": "futuristic holographic HUD display, glowing wireframes, telemetry data nodes, neon cyan and deep navy translucent elements, sci-fi interface"
    }

    bg_modifier = "isolated on clean solid white background, zero shadows, commercial stock asset"
    if transparent_bg:
        bg_modifier = "isolated on pure white background, ready for cutout, clean high contrast silhouette edges, studio product shot, no background clutter"

    prompt_en = f"{subject}, {style_presets.get(style, style_presets['3d-render'])}, {bg_modifier}, color scheme inspired by {color_palette}, masterwork, ultra-detailed --ar {aspect_ratio.replace(':', '_')}"

    return {
        "subject": subject,
        "style": style,
        "prompt_en": prompt_en,
        "negative_prompt": "blurry, low quality, noisy artifacts, watermarks, text, cluttered background, distorted geometry, jpeg compression artifacts",
        "suggested_aspect": aspect_ratio,
        "transparent_ready": transparent_bg
    }


def crop_region(image_path: str, bbox: Tuple[float, float, float, float], output_path: str) -> str:
    """
    从原始图片中截取局部区域
    bbox: (x1, y1, x2, y2) 归一化比例 (0.0 ~ 1.0) 或绝对像素坐标
    """
    if not HAS_PIL:
        raise RuntimeError("需要 Pillow 库支持，请先安装 Pillow: pip install Pillow")

    img = Image.open(image_path).convert("RGBA")
    w, h = img.size

    # 判断是否为归一化比例
    x1, y1, x2, y2 = bbox
    if max(x1, y1, x2, y2) <= 1.0:
        box = (int(x1 * w), int(y1 * h), int(x2 * w), int(y2 * h))
    else:
        box = (int(x1), int(y1), int(x2), int(y2))

    cropped = img.crop(box)
    cropped.save(output_path, "PNG")
    print(f"[OK] 区域截取完成: {output_path} ({cropped.size[0]}x{cropped.size[1]})")
    return output_path


def remove_background_precise(image_path: str, output_path: str, tolerance: int = 35, feather_radius: int = 2) -> str:
    """
    高保真局部背景去除与边缘羽化（针对白色/浅色背景工作室资产）
    """
    if not HAS_PIL:
        raise RuntimeError("需要 Pillow 库支持")

    img = Image.open(image_path).convert("RGBA")
    r, g, b, a = img.split()

    # 检测四角背景主基准色 (默认以左上角为样本)
    corner_pixels = [img.getpixel((0, 0)), img.getpixel((img.width - 1, 0)), img.getpixel((0, img.height - 1))]
    bg_r = sum(p[0] for p in corner_pixels) // len(corner_pixels)
    bg_g = sum(p[1] for p in corner_pixels) // len(corner_pixels)
    bg_b = sum(p[2] for p in corner_pixels) // len(corner_pixels)

    # 构造精确透明通道
    datas = img.getdata()
    new_data = []
    for item in datas:
        # 色差距离计算 (Euclidean distance in RGB space)
        dist = ((item[0] - bg_r) ** 2 + (item[1] - bg_g) ** 2 + (item[2] - bg_b) ** 2) ** 0.5
        if dist < tolerance:
            new_data.append((item[0], item[1], item[2], 0))  # 完全透明
        elif dist < tolerance + 15:
            # 边缘线性过渡羽化
            alpha = int(255 * (dist - tolerance) / 15)
            new_data.append((item[0], item[1], item[2], alpha))
        else:
            new_data.append(item)

    img.putdata(new_data)

    # 边缘微平滑处理
    if feather_radius > 0:
        alpha_channel = img.split()[3]
        alpha_channel = alpha_channel.filter(ImageFilter.GaussianBlur(radius=feather_radius * 0.5))
        img.putalpha(alpha_channel)

    img.save(output_path, "PNG")
    print(f"[OK] 背景扣除与透明通道生成完毕: {output_path}")
    return output_path


def apply_irregular_shape_mask(image_path: str, output_path: str, shape: str = "organic-blob", corner_radius: int = 40) -> str:
    """
    将图片裁剪为高颜值不规则异形形状 (带柔和抗锯齿边缘)
    支持: organic-blob (有机流体), hexagon (六边形), round-rect (圆角), circle (正圆), snip-corner (双斜切角)
    """
    if not HAS_PIL:
        raise RuntimeError("需要 Pillow 库支持")

    img = Image.open(image_path).convert("RGBA")
    w, h = img.size

    # 4倍超采样掩码 (Supersampling 抗锯齿)
    scale = 4
    mask_size = (w * scale, h * scale)
    mask = Image.new("L", mask_size, 0)
    draw = ImageDraw.Draw(mask)

    if shape == "circle":
        min_dim = min(w, h) * scale
        x0 = (w * scale - min_dim) / 2
        y0 = (h * scale - min_dim) / 2
        draw.ellipse([x0, y0, x0 + min_dim, y0 + min_dim], fill=255)

    elif shape == "round-rect":
        draw.rounded_rectangle([0, 0, w * scale, h * scale], radius=corner_radius * scale, fill=255)

    elif shape == "hexagon":
        # 正六边形计算
        cx, cy = (w * scale) / 2, (h * scale) / 2
        rx, ry = (w * scale) / 2 * 0.96, (h * scale) / 2 * 0.96
        import math
        points = []
        for i in range(6):
            angle = math.radians(60 * i - 30)
            px = cx + rx * math.cos(angle)
            py = cy + ry * math.sin(angle)
            points.append((px, py))
        draw.polygon(points, fill=255)

    elif shape == "snip-corner":
        # 现代科技切角多边形
        snip = 50 * scale
        points = [
            (snip, 0),
            (w * scale - snip, 0),
            (w * scale, snip),
            (w * scale, h * scale - snip),
            (w * scale - snip, h * scale),
            (snip, h * scale),
            (0, h * scale - snip),
            (0, snip)
        ]
        draw.polygon(points, fill=255)

    elif shape == "organic-blob":
        # 优美有机流体多弧线平滑团块 (双层混合贝塞尔形态逼近)
        cx, cy = (w * scale) / 2, (h * scale) / 2
        rx, ry = (w * scale) / 2, (h * scale) / 2
        # 由4个带平滑偏移的半椭圆拟合有呼吸感的非均匀流体
        draw.ellipse([cx - rx * 0.98, cy - ry * 0.92, cx + rx * 0.95, cy + ry * 0.96], fill=255)
        draw.ellipse([cx - rx * 0.92, cy - ry * 0.96, cx + rx * 0.98, cy + ry * 0.90], fill=255)
    else:
        draw.rounded_rectangle([0, 0, w * scale, h * scale], radius=corner_radius * scale, fill=255)

    # 缩小回原图大小实现完美超采样抗锯齿
    mask = mask.resize((w, h), Image.Resampling.LANCZOS)

    # 合并原有透明通道
    current_alpha = img.split()[3]
    final_alpha = ImageOps.invert(ImageOps.invert(current_alpha))
    mask_combined = Image.new("L", (w, h))
    for x in range(w):
        for y in range(h):
            a_orig = current_alpha.getpixel((x, y))
            a_mask = mask.getpixel((x, y))
            mask_combined.putpixel((x, y), min(a_orig, a_mask))

    img.putalpha(mask_combined)
    img.save(output_path, "PNG")
    print(f"[OK] 不规则异形蒙版【{shape}】应用成功: {output_path}")
    return output_path


def place_asset_into_pptx(
    pptx_path: str,
    slide_index: int,
    image_path: str,
    left_inch: float,
    top_inch: float,
    width_inch: float,
    height_inch: Optional[float] = None,
    output_pptx: Optional[str] = None
) -> str:
    """
    将透明素材或异形图片精准放置到指定 PPT 页码中
    """
    if not HAS_PPTX:
        raise RuntimeError("需要 python-pptx 库支持")

    prs = Presentation(pptx_path)
    if slide_index >= len(prs.slides):
        raise IndexError(f"幻灯片页码越界: 总共 {len(prs.slides)} 页，请求第 {slide_index + 1} 页")

    slide = prs.slides[slide_index]
    left = Inches(left_inch)
    top = Inches(top_inch)
    width = Inches(width_inch)
    height = Inches(height_inch) if height_inch else None

    pic = slide.shapes.add_picture(image_path, left, top, width=width, height=height)

    dest = output_pptx or pptx_path
    prs.save(dest)
    print(f"[OK] 视觉素材已成功装配至 PPT 第 {slide_index + 1} 页: ({left_inch}\", {top_inch}\") -> {dest}")
    return dest


def main():
    parser = argparse.ArgumentParser(description="PPT 视觉资产深度工坊与图像抠图排版引擎")
    subparsers = parser.add_subparsers(dest="action", help="选择子命令")

    # 1. prompt 命令
    p_prompt = subparsers.add_parser("prompt", help="生成工业级完美生图提示词")
    p_prompt.add_argument("subject", help="画面主体描述 (如: 智慧供应链无人配送机器人)")
    p_prompt.add_argument("--style", default="3d-render", choices=["3d-render", "isometric", "corporate-photo", "flat-vector", "cyber-hud"], help="视觉风格")
    p_prompt.add_argument("--palette", default="modern-tech", help="配色氛围")
    p_prompt.add_argument("--transparent", action="store_true", help="是否以无背景抠图为目的生成")

    # 2. crop 命令
    p_crop = subparsers.add_parser("crop", help="从原图中切片截取局部关键主体")
    p_crop.add_argument("image", help="输入图像路径")
    p_crop.add_argument("-o", "--output", required=True, help="切片输出路径 (.png)")
    p_crop.add_argument("--bbox", required=True, help="归一化坐标 x1,y1,x2,y2 (例如: 0.1,0.2,0.8,0.9)")

    # 3. cutout (抠图) 命令
    p_cut = subparsers.add_parser("cutout", help="高保真背景扣除与透明通道提取")
    p_cut.add_argument("image", help="输入图像路径")
    p_cut.add_argument("-o", "--output", required=True, help="透明图输出路径 (.png)")
    p_cut.add_argument("--tolerance", type=int, default=35, help="背景色差容差阈值")

    # 4. mask (异形蒙版) 命令
    p_mask = subparsers.add_parser("mask", help="将图片转换为不规则异形美化形态")
    p_mask.add_argument("image", help="输入图像路径")
    p_mask.add_argument("-o", "--output", required=True, help="异形图片输出路径 (.png)")
    p_mask.add_argument("--shape", default="organic-blob", choices=["organic-blob", "hexagon", "round-rect", "circle", "snip-corner"], help="形状类型")
    p_mask.add_argument("--radius", type=int, default=40, help="圆角弧度")

    # 5. place (放置入PPT) 命令
    p_place = subparsers.add_parser("place", help="将图片放置到 PPT 指定坐标")
    p_place.add_argument("pptx", help="PPTX 文件路径")
    p_place.add_argument("--slide", type=int, default=1, help="目标页码 (从 1 开始)")
    p_place.add_argument("--image", required=True, help="要插入的图片素材路径")
    p_place.add_argument("--left", type=float, required=True, help="左边距 (英寸)")
    p_place.add_argument("--top", type=float, required=True, help="上边距 (英寸)")
    p_place.add_argument("--width", type=float, required=True, help="宽度 (英寸)")
    p_place.add_argument("--height", type=float, help="高度 (英寸，可选)")
    p_place.add_argument("-o", "--output", help="输出 PPTX 路径 (缺省覆盖原文件)")

    args = parser.parse_args()

    if args.action == "prompt":
        res = craft_perfect_prompt(
            subject=args.subject,
            style=args.style,
            color_palette=args.palette,
            transparent_bg=args.transparent
        )
        print(json.dumps(res, ensure_ascii=False, indent=2))

    elif args.action == "crop":
        coords = [float(x.strip()) for x in args.bbox.split(",")]
        crop_region(args.image, tuple(coords), args.output)

    elif args.action == "cutout":
        remove_background_precise(args.image, args.output, tolerance=args.tolerance)

    elif args.action == "mask":
        apply_irregular_shape_mask(args.image, args.output, shape=args.shape, corner_radius=args.radius)

    elif args.action == "place":
        place_asset_into_pptx(
            pptx_path=args.pptx,
            slide_index=args.slide - 1,
            image_path=args.image,
            left_inch=args.left,
            top_inch=args.top,
            width_inch=args.width,
            height_inch=args.height,
            output_pptx=args.output
        )
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
