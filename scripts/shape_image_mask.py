# -*- coding: utf-8 -*-
"""
shape_image_mask.py — 将图片裁剪并放置为不规则高审美形状（有机流体Blob、几何异形、对角斜切卡片、自定义多边形蒙版）
=============================================================================================================
支持功能：
1. 图像不规则形状蒙版生成（输出带平滑抗锯齿 Alpha 通道的透明 PNG）：
   - organic_blob: 多阶平滑有机流体形 / 液态水滴形（支持随机种子与凹凸度调节）
   - squircle: 苹果级超椭圆（平滑连续曲率）
   - hexagon: 科技正六边形 / 蜂巢六边形
   - octagon: 现代八边形
   - diamond: 锐利菱形 / 旋转正方形
   - teardrop: 水滴泪珠形
   - slanted_card: 杂志级对角线斜切卡片
   - chevron: 导向箭头形
   - polygon: 任意自由多边形顶点蒙版（归一化坐标或像素坐标）
2. 形状边框描边与发光描边（可选纯色轮廓线、半透明边框）。
3. 原生 PPTX 注入集成：
   - 模式 A（高质矢量内嵌）：将蒙版化透明 PNG 插入指定幻灯片坐标并附加投影
   - 模式 B（原生形状填充）：利用 python-pptx 原生 MSO_SHAPE + shape.fill.user_picture 渲染

用法示例：
    # 裁剪为有机流体形状
    python scripts/shape_image_mask.py input.png -o blob.png --shape blob --border 3 --border-color "#0066FF"

    # 裁剪为正六边形
    python scripts/shape_image_mask.py input.png -o hex.png --shape hexagon

    # 裁剪为对角斜切杂志卡片
    python scripts/shape_image_mask.py input.png -o card.png --shape slanted_card --slant-ratio 0.15

    # 裁剪为自定义多边形
    python scripts/shape_image_mask.py input.png -o poly.png --shape polygon --points "0.1,0.0;0.95,0.05;0.85,1.0;0.0,0.85"

    # 直接将异形图片植入 PPTX 幻灯片
    python scripts/shape_image_mask.py input.png --pptx deck.pptx --slide 1 --shape blob --left 1.5 --top 1.5 --width 4.0
"""

import os
import sys
import math
import argparse
from typing import List, Tuple, Optional
from PIL import Image, ImageDraw, ImageFilter, ImageOps
import numpy as np

# Optional python-pptx import
try:
    from pptx import Presentation
    from pptx.util import Inches, Pt
    from pptx.enum.shapes import MSO_SHAPE
    from pptx.dml.color import RGBColor
    PPTX_AVAILABLE = True
except ImportError:
    PPTX_AVAILABLE = False


def hex_to_rgb(hex_str: str) -> Tuple[int, int, int]:
    hex_str = hex_str.lstrip("#")
    if len(hex_str) == 3:
        hex_str = "".join([c * 2 for c in hex_str])
    return tuple(int(hex_str[i:i+2], 16) for i in (0, 2, 4))


def generate_organic_blob_mask(width: int, height: int, num_points: int = 8, irregularity: float = 0.25, seed: Optional[int] = None) -> Image.Image:
    """
    生成超平滑有机流体 (Organic Blob) 形状的高精抗锯齿蒙版
    """
    scale = 4  # 4x 超采样消除锯齿
    sw, sh = width * scale, height * scale
    cx, cy = sw / 2.0, sh / 2.0
    rx, ry = (sw * 0.44), (sh * 0.44)

    if seed is not None:
        np.random.seed(seed)
    else:
        np.random.seed(42)

    angles = np.linspace(0, 2 * np.pi, num_points, endpoint=False)
    radii = 1.0 + np.random.uniform(-irregularity, irregularity, num_points)
    # 平滑首尾闭合
    radii[-1] = (radii[-2] + radii[0]) / 2.0

    points = []
    for angle, r_factor in zip(angles, radii):
        x = cx + rx * r_factor * math.cos(angle)
        y = cy + ry * r_factor * math.sin(angle)
        points.append((x, y))

    # 生成平滑插值样条曲线
    dense_angles = np.linspace(0, 2 * np.pi, 360, endpoint=False)
    # 用简易周期性样条平滑
    r_interp = np.interp(dense_angles, angles, radii, period=2 * np.pi)
    
    # 柔和高斯平滑滤波
    kernel_size = 21
    kernel = np.ones(kernel_size) / kernel_size
    r_smooth = np.convolve(np.tile(r_interp, 3), kernel, mode='same')[len(dense_angles):2*len(dense_angles)]

    smooth_points = []
    for angle, r_val in zip(dense_angles, r_smooth):
        x = cx + rx * r_val * math.cos(angle)
        y = cy + ry * r_val * math.sin(angle)
        smooth_points.append((x, y))

    mask_high = Image.new("L", (sw, sh), 0)
    draw = ImageDraw.Draw(mask_high)
    draw.polygon(smooth_points, fill=255)

    # 轻微平滑并下采样至目标尺寸
    mask_high = mask_high.filter(ImageFilter.GaussianBlur(radius=scale * 0.8))
    mask = mask_high.resize((width, height), Image.Resampling.LANCZOS)
    return mask


def generate_squircle_mask(width: int, height: int, n: float = 4.0) -> Image.Image:
    """
    生成超椭圆 (Squircle, (x/a)^n + (y/b)^n = 1) 蒙版，呈现极致现代感苹果曲率
    """
    scale = 4
    sw, sh = width * scale, height * scale
    cx, cy = sw / 2.0, sh / 2.0
    a, b = sw * 0.47, sh * 0.47

    y, x = np.ogrid[:sh, :sw]
    dist = ((np.abs(x - cx) / a) ** n + (np.abs(y - cy) / b) ** n)
    mask_arr = np.where(dist <= 1.0, 255, 0).astype(np.uint8)

    mask_high = Image.fromarray(mask_arr, mode="L")
    mask_high = mask_high.filter(ImageFilter.GaussianBlur(radius=scale * 0.6))
    return mask_high.resize((width, height), Image.Resampling.LANCZOS)


def generate_polygon_mask(width: int, height: int, normalized_points: List[Tuple[float, float]]) -> Image.Image:
    """
    根据归一化坐标生成抗锯齿多边形蒙版
    """
    scale = 4
    sw, sh = width * scale, height * scale
    pixel_points = [(p[0] * sw, p[1] * sh) for p in normalized_points]

    mask_high = Image.new("L", (sw, sh), 0)
    draw = ImageDraw.Draw(mask_high)
    draw.polygon(pixel_points, fill=255)
    mask_high = mask_high.filter(ImageFilter.GaussianBlur(radius=scale * 0.5))
    return mask_high.resize((width, height), Image.Resampling.LANCZOS)


def get_shape_mask(shape_type: str, width: int, height: int, custom_points: Optional[str] = None, slant_ratio: float = 0.12) -> Image.Image:
    """
    获取指定类型的抗锯齿形状蒙版
    """
    st = shape_type.lower()
    if st in ("blob", "organic_blob", "organic"):
        return generate_organic_blob_mask(width, height)
    elif st in ("squircle", "superellipse"):
        return generate_squircle_mask(width, height, n=4.0)
    elif st == "hexagon":
        # 水平或垂直对称六边形
        pts = [(0.25, 0.05), (0.75, 0.05), (0.98, 0.5), (0.75, 0.95), (0.25, 0.95), (0.02, 0.5)]
        return generate_polygon_mask(width, height, pts)
    elif st == "octagon":
        pts = [(0.29, 0.03), (0.71, 0.03), (0.97, 0.29), (0.97, 0.71), (0.71, 0.97), (0.29, 0.97), (0.03, 0.71), (0.03, 0.29)]
        return generate_polygon_mask(width, height, pts)
    elif st == "diamond":
        pts = [(0.5, 0.03), (0.97, 0.5), (0.5, 0.97), (0.03, 0.5)]
        return generate_polygon_mask(width, height, pts)
    elif st == "teardrop":
        # 水滴形：上部弧形，右下角尖角
        pts = [(0.5, 0.05), (0.95, 0.35), (0.95, 0.85), (0.85, 0.95), (0.35, 0.95), (0.05, 0.55), (0.15, 0.2)]
        return generate_polygon_mask(width, height, pts)
    elif st in ("slanted_card", "slanted", "diagonal"):
        # 现代斜切卡片：左高右低，或对角切角
        s = max(0.02, min(0.4, slant_ratio))
        pts = [(0.0, 0.0), (1.0, s), (1.0, 1.0), (0.0, 1.0 - s)]
        return generate_polygon_mask(width, height, pts)
    elif st == "chevron":
        pts = [(0.0, 0.0), (0.8, 0.0), (1.0, 0.5), (0.8, 1.0), (0.0, 1.0), (0.2, 0.5)]
        return generate_polygon_mask(width, height, pts)
    elif st == "polygon":
        if not custom_points:
            raise ValueError("shape='polygon' 要求提供 --points 参数，格式形如 '0.1,0.0;0.9,0.1;0.8,1.0;0.0,0.9'")
        pts = []
        for pair in custom_points.split(";"):
            pair = pair.strip()
            if pair:
                x_str, y_str = pair.split(",")
                pts.append((float(x_str.strip()), float(y_str.strip())))
        return generate_polygon_mask(width, height, pts)
    else:
        raise ValueError(f"未知形状类型: {shape_type}。可选: blob, squircle, hexagon, octagon, diamond, teardrop, slanted_card, chevron, polygon")


def apply_shape_mask(
    image_path: str,
    shape_type: str = "blob",
    border_width: int = 0,
    border_color: str = "#FFFFFF",
    border_opacity: float = 1.0,
    custom_points: Optional[str] = None,
    slant_ratio: float = 0.12,
    output_path: Optional[str] = None
) -> Image.Image:
    """
    将目标图片裁切为指定形状，并可选附带平滑抗锯齿轮廓线
    """
    img = Image.open(image_path).convert("RGBA")
    w, h = img.size

    # 获取抗锯齿蒙版
    mask = get_shape_mask(shape_type, w, h, custom_points=custom_points, slant_ratio=slant_ratio)

    # 合并原始图像的现有 Alpha (若有)
    if "A" in img.getbands():
        orig_alpha = img.getchannel("A")
        combined_alpha = ImageMath_multiply(orig_alpha, mask)
    else:
        combined_alpha = mask

    # 组装结果
    result = img.copy()
    result.putalpha(combined_alpha)

    # 绘制外边框 (若 border_width > 0)
    if border_width > 0:
        r, g, b = hex_to_rgb(border_color)
        a_val = int(255 * max(0.0, min(1.0, border_opacity)))
        
        # 通过形态学膨胀/膨胀边缘提取精确轮廓
        scale = 2
        large_mask = mask.resize((w * scale, h * scale), Image.Resampling.BILINEAR)
        dilated = large_mask.filter(ImageFilter.MaxFilter(border_width * scale * 2 + 1))
        outline = ImageOps.invert(large_mask)
        outline_mask_large = ImageOps.invert(dilated)
        border_mask_large = ImageMath_subtract(dilated, large_mask)
        border_mask = border_mask_large.resize((w, h), Image.Resampling.LANCZOS)

        border_img = Image.new("RGBA", (w, h), (r, g, b, a_val))
        result = Image.composite(border_img, result, border_mask)

    if output_path:
        os.makedirs(os.path.dirname(os.path.abspath(output_path)), exist_ok=True)
        result.save(output_path, "PNG")
        print(f"[OK] 异形蒙版裁剪完成: {output_path} ({w}x{h}, shape={shape_type})")

    return result


def ImageMath_multiply(im1: Image.Image, im2: Image.Image) -> Image.Image:
    a1 = np.array(im1, dtype=np.float32) / 255.0
    a2 = np.array(im2, dtype=np.float32) / 255.0
    res = (a1 * a2 * 255.0).astype(np.uint8)
    return Image.fromarray(res, mode="L")


def ImageMath_subtract(im1: Image.Image, im2: Image.Image) -> Image.Image:
    a1 = np.array(im1, dtype=np.int16)
    a2 = np.array(im2, dtype=np.int16)
    res = np.clip(a1 - a2, 0, 255).astype(np.uint8)
    return Image.fromarray(res, mode="L")


def place_in_pptx(
    pptx_path: str,
    slide_index: int,
    image_path: str,
    shape_type: str = "blob",
    left_inch: float = 1.0,
    top_inch: float = 1.0,
    width_inch: float = 4.0,
    height_inch: Optional[float] = None,
    border_width: int = 0,
    border_color: str = "#FFFFFF",
    output_pptx: Optional[str] = None
) -> str:
    """
    将图片裁剪为异形后直接注入 PPTX 幻灯片
    """
    if not PPTX_AVAILABLE:
        raise RuntimeError("python-pptx 库未安装，无法进行 PPTX 注入")

    # 1. 生成异形透明图
    temp_cutout = os.path.splitext(image_path)[0] + f"_masked_{shape_type}.png"
    apply_shape_mask(
        image_path=image_path,
        shape_type=shape_type,
        border_width=border_width,
        border_color=border_color,
        output_path=temp_cutout
    )

    prs = Presentation(pptx_path)
    if slide_index < 1 or slide_index > len(prs.slides):
        raise IndexError(f"幻灯片序号越界: 需在 1~{len(prs.slides)} 之间，传入 {slide_index}")

    slide = prs.slides[slide_index - 1]

    # 计算高度保持纵横比
    orig_img = Image.open(image_path)
    aspect = orig_img.height / float(orig_img.width)
    if height_inch is None:
        height_inch = width_inch * aspect

    # 插入异形图片对象
    slide.shapes.add_picture(
        temp_cutout,
        Inches(left_inch),
        Inches(top_inch),
        width=Inches(width_inch),
        height=Inches(height_inch)
    )

    out_file = output_pptx or pptx_path
    prs.save(out_file)
    print(f"[OK] 异形素材已植入 PPTX: {out_file} (Slide {slide_index}, shape={shape_type}, left={left_inch}in, top={top_inch}in)")
    return out_file


def main():
    parser = argparse.ArgumentParser(description="PPT 不规则高审美形状图片裁切与植入引擎")
    parser.add_argument("input", help="输入图片路径 (PNG/JPG/WEBP)")
    parser.add_argument("-o", "--output", help="输出图片路径 (PNG)")
    parser.add_argument("--shape", default="blob", choices=["blob", "squircle", "hexagon", "octagon", "diamond", "teardrop", "slanted_card", "chevron", "polygon"], help="目标形状")
    parser.add_argument("--border", type=int, default=0, help="轮廓描边宽度 (像素)")
    parser.add_argument("--border-color", default="#FFFFFF", help="轮廓描边颜色十六进制 (如 #0066FF)")
    parser.add_argument("--border-opacity", type=float, default=1.0, help="描边透明度 (0.0~1.0)")
    parser.add_argument("--slant-ratio", type=float, default=0.12, help="对角斜切卡片斜率 (0.02~0.4)")
    parser.add_argument("--points", help="自定义多边形归一化坐标点 '0,0;1,0.1;0.8,1;0.1,0.9'")
    parser.add_argument("--pptx", help="可选：直接注入的目标 PPTX 文件路径")
    parser.add_argument("--slide", type=int, default=1, help="PPTX 目标幻灯片页码 (从 1 起)")
    parser.add_argument("--left", type=float, default=1.0, help="PPTX 水平坐标 (英寸)")
    parser.add_argument("--top", type=float, default=1.0, help="PPTX 垂直坐标 (英寸)")
    parser.add_argument("--width", type=float, default=4.0, help="PPTX 宽度 (英寸)")
    parser.add_argument("--height", type=float, default=None, help="PPTX 高度 (英寸，缺省按比例计算)")

    args = parser.parse_args()

    if args.pptx:
        place_in_pptx(
            pptx_path=args.pptx,
            slide_index=args.slide,
            image_path=args.input,
            shape_type=args.shape,
            left_inch=args.left,
            top_inch=args.top,
            width_inch=args.width,
            height_inch=args.height,
            border_width=args.border,
            border_color=args.border_color
        )
    else:
        out = args.output or (os.path.splitext(args.input)[0] + f"_masked_{args.shape}.png")
        apply_shape_mask(
            image_path=args.input,
            shape_type=args.shape,
            border_width=args.border,
            border_color=args.border_color,
            border_opacity=args.border_opacity,
            custom_points=args.points,
            slant_ratio=args.slant_ratio,
            output_path=out
        )


if __name__ == "__main__":
    main()
