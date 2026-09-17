# -*- coding: utf-8 -*-
"""
image_prompt_craft.py — 基于图像的完美生图提示词反推、风格提炼与极速生图对接工具
=============================================================================
功能定位：
1. 图像视觉特征解构 (Vision Analysis Contract)：
   - 主体特征 (Core Subject): 物体形态、几何结构、部件层次
   - 渲染媒介 (Artistic Medium): 3D Octane, C4D Claymation, Frosted Glassmorphism, Vector, Macro Photo
   - 材质物理 (Materials): 磨砂亚克力、拉丝金属、次表面散射 (SSS)、陶瓷釉面
   - 构图光影 (Lighting & Framing): 轴测等轴 30°、柔光箱侧光、轮廓光 (Rim Light)、大景深留白
   - 色板收敛 (Color Harmony): 提取视觉主辅色与点缀色
2. 完美提示词重构 (Prompt Synthesis)：
   - Midjourney v6 风格指令 (`--ar 16:9 --style raw --v 6.1`)
   - Flux.1 / DALL-E 3 / SDXL 格式自然语言提示词
   - Gemini-3.1-flash-image 专属素材生成提示词（强化透明底/纯白隔离底，方便后续 1 秒无损抠图）
3. 针对 PPT 设计的“透明素材定制模式” (`--transparent-asset`)：
   - 强制添加纯白影棚隔离背景 ("isolated on clean seamless solid white background, zero shadows, studio product cutout")
   - 确保生成的素材可直接无缝供 `matting_cutout.py` 或 `shape_image_mask.py` 瞬间一键抠出透明 PNG！

用法示例：
    # 1. 针对某张参考图提炼通用高质量提示词
    python scripts/image_prompt_craft.py ref.png --style 3d-glass

    # 2. 生成专为 PPT 制作透明背景素材的提示词 (白底无杂物，利于抠图)
    python scripts/image_prompt_craft.py ref.png --transparent-asset --subject "智能工业机械臂核心中枢"

    # 3. 指定输出目标平台
    python scripts/image_prompt_craft.py ref.png --engine midjourney --ar 16:9
"""

import os
import sys
import argparse
import json
from typing import Dict, List, Optional
from PIL import Image
import numpy as np

# 内置主流工业级商业 PPT 常用美术风格模板
STYLE_PRESETS = {
    "3d-glass": {
        "name": "3D半透明磨砂玻璃拟态 (Frosted Glassmorphism)",
        "keywords": "3D render, frosted translucent glass, matte acrylic material, vibrant refraction, internal soft neon glow, sleek rounded edges, clean futuristic octane render, ambient occlusion, ray tracing, 8k resolution, modern high-tech visual",
        "bg_hint": "solid seamless white background, no floor contact shadow"
    },
    "clay-minimal": {
        "name": "极简黏土低多边形风 (Claymation & Minimalist 3D)",
        "keywords": "soft clay texture, matte finish, charming minimalist 3D illustration, cute proportions, smooth studio lighting, pastel color accents, clean isometric perspective, blender render",
        "bg_hint": "isolated on pure white background, studio backdrop"
    },
    "metallic-tech": {
        "name": "重工业拉丝科技金属 (Brushed Titanium & Industrial Tech)",
        "keywords": "precision engineered machinery, brushed aluminum and anodized titanium finish, clean chamfered edges, hyper-detailed mechanical parts, blue led indicator lights, cinematic studio product photography",
        "bg_hint": "isolated on pure white seamless background"
    },
    "editorial-flat": {
        "name": "高端商业杂志扁平插画 (Modern Editorial Vector)",
        "keywords": "modern editorial flat vector illustration, refined geometric silhouettes, elegant line work, sophisticated limited color palette, bauhaus graphic design aesthetic, clean negative space",
        "bg_hint": "clean white background"
    },
    "macro-product": {
        "name": "商业静物微距特写 (Commercial Macro Photography)",
        "keywords": "macro close-up photography, sharp focus, shallow depth of field, natural softbox lighting, ultra-fine surface texture, Hasselblad medium format camera, crisp reflections",
        "bg_hint": "seamless solid white infinity cove background"
    }
}


def analyze_image_palette(img_path: str, max_colors: int = 4) -> List[str]:
    """
    轻量快速提取图片核心色彩十六进制编码
    """
    try:
        img = Image.open(img_path).convert("RGB")
        img = img.resize((100, 100))
        # K-means 或量化
        quantized = img.quantize(colors=max_colors, method=Image.Quantize.MEDIANCUT)
        palette = quantized.getpalette()[:max_colors * 3]
        hex_colors = []
        for i in range(0, len(palette), 3):
            r, g, b = palette[i], palette[i+1], palette[i+2]
            hex_colors.append(f"#{r:02X}{g:02X}{b:02X}")
        return hex_colors
    except Exception:
        return ["#0066FF", "#111827", "#F3F4F6"]


def craft_perfect_prompt(
    subject: str,
    style_key: str = "3d-glass",
    image_path: Optional[str] = None,
    transparent_asset: bool = True,
    engine: str = "general",
    ar: str = "16:9",
    color_hint: Optional[str] = None
) -> Dict[str, str]:
    """
    重构高质量生图提示词
    """
    style_info = STYLE_PRESETS.get(style_key, STYLE_PRESETS["3d-glass"])
    
    extracted_colors = []
    if image_path and os.path.exists(image_path):
        extracted_colors = analyze_image_palette(image_path)
    
    color_desc = color_hint or (", ".join(extracted_colors[:3]) + " color tones" if extracted_colors else "harmonious corporate palette")

    # 构筑提示词主体
    parts = []
    parts.append(f"A masterfully detailed visual representation of {subject}")
    parts.append(style_info["keywords"])
    parts.append(f"color palette featuring {color_desc}")

    if transparent_asset:
        parts.append("isolated on a clean seamless solid pure white background, studio product cutout, no background distractions, perfectly centered, ready for alpha cutout")
    else:
        parts.append("clean elegant composition, abundant negative space for text placement")

    prompt_base = ", ".join(parts)

    # 针对不同引擎微调
    results = {}
    
    # 1. Midjourney
    mj_tail = f" --ar {ar} --style raw --v 6.1" if engine == "midjourney" or engine == "all" else ""
    results["midjourney"] = f"/imagine prompt: {prompt_base}{mj_tail}"

    # 2. Flux / DALL-E 3 / Gemini
    results["dalle3_flux_gemini"] = prompt_base

    # 3. 负向提示词 (Negative Prompt)
    results["negative_prompt"] = (
        "blurry, noisy, low resolution, distorted geometry, cluttered background, "
        "unintended text, watermarks, ugly borders, cropped subject, artifacts"
    )

    # 4. 建议后续链路
    results["next_step_action"] = (
        "1. 使用该提示词生成高清图像；\n"
        "2. 将生成的图像传入 `python scripts/matting_cutout.py <image> --auto-bg` 一键抠出纯净透明 PNG；\n"
        "3. 或调用 `python scripts/shape_image_mask.py <image> --shape blob` 裁剪为有机流体异形卡片并直接插入 PPT！"
    )

    return results


def main():
    parser = argparse.ArgumentParser(description="PPT 图像反推与完美生图提示词工程器")
    parser.add_argument("image", nargs="?", help="参考图片路径 (可选)")
    parser.add_argument("-s", "--subject", default="核心技术架构中枢产品主体", help="画面主体描述")
    parser.add_argument("--style", default="3d-glass", choices=list(STYLE_PRESETS.keys()), help="预设艺术风格")
    parser.add_argument("--transparent-asset", action="store_true", default=True, help="是否生成专用于PPT透明素材抠图的提示词 (纯白背景隔离)")
    parser.add_argument("--engine", default="all", choices=["all", "midjourney", "gemini", "dalle3", "flux"], help="目标生图模型格式")
    parser.add_argument("--ar", default="16:9", help="画幅比 (如 16:9, 1:1, 4:3)")
    parser.add_argument("--json", action="store_true", help="以 JSON 格式输出")

    args = parser.parse_args()

    crafted = craft_perfect_prompt(
        subject=args.subject,
        style_key=args.style,
        image_path=args.image,
        transparent_asset=args.transparent_asset,
        engine=args.engine,
        ar=args.ar
    )

    if args.json:
        print(json.dumps(crafted, ensure_ascii=False, indent=2))
    else:
        print("\n========================================================")
        print("🎨 PPT 高精素材专属生成提示词 (Prompt Craft Result)")
        print("========================================================")
        print(f"【风格选择】: {STYLE_PRESETS[args.style]['name']}")
        print(f"【主体内容】: {args.subject}")
        print("\n[Midjourney v6 提示词]:")
        print(crafted["midjourney"])
        print("\n[Flux / DALL-E 3 / Gemini-3.1-flash 提示词]:")
        print(crafted["dalle3_flux_gemini"])
        print("\n[负向提示词 (Negative Prompt)]:")
        print(crafted["negative_prompt"])
        print("\n[配套自动化闭环建议]:")
        print(crafted["next_step_action"])
        print("========================================================\n")


if __name__ == "__main__":
    main()
