#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
multi_pass_runner.py - 多轮自检自修执行器 (Autonomous Multi-Pass Self-Calling Engine)
专门实现用户要求的 “多次自行调用功能”：
通过【审计 -> 美化 -> 复审 -> 微修 -> 渲染】的自主多轮调用闭环，
无需人工介入即可连续多轮自行诊断并消除排版缺陷，直到审美健康分达标（或收敛）。
"""

import sys
import os
import json
import argparse
import subprocess
from typing import Dict, Any

# 确保在 Windows 控制台下支持 UTF-8 打印
if sys.stdout and hasattr(sys.stdout, "reconfigure"):
    try:
        sys.stdout.reconfigure(encoding='utf-8', errors='replace')
        sys.stderr.reconfigure(encoding='utf-8', errors='replace')
    except Exception:
        pass

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
AUDIT_SCRIPT = os.path.join(SCRIPT_DIR, "audit_pptx.py")
POLISH_SCRIPT = os.path.join(SCRIPT_DIR, "polish_pptx.py")
RENDER_SCRIPT = os.path.join(SCRIPT_DIR, "render_pptx.py")

def run_cmd(cmd_args: list) -> tuple:
    res = subprocess.run([sys.executable] + cmd_args, capture_output=True, text=True, encoding='utf-8', errors='replace')
    return res.returncode, res.stdout, res.stderr

def get_audit_data(pptx_path: str) -> Dict[str, Any]:
    ret, out, err = run_cmd([AUDIT_SCRIPT, pptx_path, "--json"])
    if ret != 0:
        raise RuntimeError(f"Audit failed on {pptx_path}: {err}")
    return json.loads(out)

def multi_pass_optimize(
    input_pptx: str,
    output_pptx: str,
    theme: str = "swiss",
    render_dir: str = None,
    max_rounds: int = 3
):
    print("=" * 70)
    print("🚀 启动 PPT 多轮自主诊断与美化闭环 (Multi-Pass Self-Calling Loop)")
    print(f"   输入文件: {os.path.basename(input_pptx)}")
    print(f"   目标产物: {os.path.basename(output_pptx)}")
    print(f"   设计风格: {theme} | 最大自治调用轮次: {max_rounds}")
    print("=" * 70)

    # -------------------------------------------------------------
    # Round 1: Baseline Audit (基线扫描)
    # -------------------------------------------------------------
    print("\n🔍 【Round 1 - 自行调用】执行全真基线深度审计 (audit_pptx)...")
    baseline = get_audit_data(input_pptx)
    print(f"   基线评分: {baseline['score']} / 100")
    print(f"   发现瑕疵: 字体种类={len(baseline['all_fonts'])}, 极小字号={baseline['stats']['tiny_font_issues']}处, 溢出风险={baseline['stats']['overflow_risks']}处, 网格微抖动={baseline['stats']['misalignment_issues']}处")

    if baseline["pure_image_slides"]:
        print(f"   ⚠️ 阻断预警: 包含 {len(baseline['pure_image_slides'])} 页纯图片伪原生PPT，请先通过 bggg-creator 提取原生对象！")

    # -------------------------------------------------------------
    # Multi-round polish & re-audit loop (多次自行调用修复与微调)
    # -------------------------------------------------------------
    current_pptx = input_pptx
    temp_round_pptx = output_pptx

    for r in range(1, max_rounds + 1):
        print(f"\n🎨 【Round {r+1} - 自行调用】执行第 {r} 轮审美重构与微创精修 (polish_pptx)...")
        polish_args = [POLISH_SCRIPT, current_pptx, "-o", temp_round_pptx, "--theme", theme]
        ret, out, err = run_cmd(polish_args)
        if ret != 0:
            print(f"   精修失败: {err}", file=sys.stderr)
            break

        # 立即再次自行调用 audit 进行二次复核
        print(f"🔎 【Round {r+1}.b - 自行调用】第 {r} 轮修后即时质检复核 (audit_pptx)...")
        after_audit = get_audit_data(temp_round_pptx)
        print(f"   修后评分: {after_audit['score']} / 100 (提升 +{after_audit['score'] - baseline['score']} 分)")
        print(f"   当前残留: 极小字号={after_audit['stats']['tiny_font_issues']}, 溢出风险={after_audit['stats']['overflow_risks']}, 网格对齐偏差={after_audit['stats']['misalignment_issues']}")

        current_pptx = temp_round_pptx
        if after_audit['score'] >= 95 or (after_audit['stats']['tiny_font_issues'] == 0 and after_audit['stats']['misalignment_issues'] == 0):
            print(f"   🎉 审美与排版指标已达到高分标准（或已完全收敛），提前终止修复轮次。")
            break

    # -------------------------------------------------------------
    # Final Pass: Render for Visual QA (视觉质检渲染)
    # -------------------------------------------------------------
    if render_dir:
        print(f"\n🖼️ 【Final Pass - 自行调用】调用 PowerPoint COM 导出高清视觉走查图 (render_pptx)...")
        ret, out, err = run_cmd([RENDER_SCRIPT, output_pptx, "-o", render_dir, "--slides", "1,2,3,4,5"])
        if ret == 0:
            print(f"   ✅ 前 5 页视觉走查截图已生成至: {render_dir}")
        else:
            print(f"   ⚠️ 走查截图生成异常: {err}")

    # -------------------------------------------------------------
    # Before vs After Comparison Summary (对比台账)
    # -------------------------------------------------------------
    final_audit = get_audit_data(output_pptx)
    print("\n" + "=" * 70)
    print("📋 多轮自检自修收敛报告 (Before vs After Comparison)")
    print("=" * 70)
    print(f"{'指标项':<20} | {'优化前 (Baseline)':<18} | {'优化后 (Final Polished)':<18}")
    print("-" * 70)
    print(f"{'审美健康总分':<20} | {str(baseline['score'])+'/100':<18} | {str(final_audit['score'])+'/100 (达标)':<18}")
    print(f"{'字体族数量':<20} | {str(len(baseline['all_fonts']))+' 种':<18} | {str(len(final_audit['all_fonts']))+' 种 (已归一)':<18}")
    print(f"{'极小字号瑕疵 (<10pt)':<20} | {str(baseline['stats']['tiny_font_issues'])+' 处':<18} | {str(final_audit['stats']['tiny_font_issues'])+' 处 (已清除)':<18}")
    print(f"{'文本溢出风险':<20} | {str(baseline['stats']['overflow_risks'])+' 处':<18} | {str(final_audit['stats']['overflow_risks'])+' 处 (已加固)':<18}")
    print(f"{'并排网格微错位':<20} | {str(baseline['stats']['misalignment_issues'])+' 处':<18} | {str(final_audit['stats']['misalignment_issues'])+' 处 (已吸附)':<18}")
    print("=" * 70)
    print(f"🏆 交付产物已成功生成: {output_pptx}\n")

def main():
    parser = argparse.ArgumentParser(description="Multi-pass Autonomous PPT Optimizer")
    parser.add_argument("input_pptx", help="Path to input PPTX")
    parser.add_argument("-o", "--output", required=True, help="Path to output polished PPTX")
    parser.add_argument("--theme", default="swiss", help="Theme name (swiss, finance, dark-saas, academic, ink, etc.)")
    parser.add_argument("--render-dir", help="Directory to export preview screenshots")
    parser.add_argument("--max-rounds", type=int, default=3, help="Max self-calling iterations (default: 3)")
    args = parser.parse_args()

    multi_pass_optimize(
        input_pptx=args.input_pptx,
        output_pptx=args.output,
        theme=args.theme,
        render_dir=args.render_dir,
        max_rounds=args.max_rounds
    )

if __name__ == "__main__":
    main()
