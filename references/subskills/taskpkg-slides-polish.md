# taskpkg-slides-polish: 存量 PPTX 逐页审计与微创精修规范

> 对应顶流开源规范 `wanshuiyin/slides-polish`。
> 核心定位：**存量 PPT 逐页质检（QA）+ 微创精细化润色**。适合底子尚可、逻辑框架已定，但存在排版毛刺、文字溢出、字体字号失范、对比度不足的演示文稿。**只做微创精修，不推翻原有叙事逻辑**。

---

## 1. 核心定位与原则

- **微创原则 (Minimally Invasive)**：保持原稿的结构布局与信息传达意图，不随意增删卡片或大改布局。
- **逐页扫描 (Slide-by-Slide Audit)**：对全篇每一页的所有 Shape、TextFrame、Table、Chart 进行多维量化指标检查。
- **可度量收敛 (Measurable Convergence)**：优化前出具基线报告，优化后即时复审对比，确保各项指标实质提升。

---

## 2. 核心质检指标体系 (The 6 Audit Dimensions)

| 检查维度 | 判定红线 | 严重等级 | 修复策略 |
|---|---|---|---|
| **纯图片扁平页** | 单张图片覆盖率 ≥ 75% 且无原生文本 | Blocker (阻断) | 必须使用 `bggg-creator` 重新拆解为原生对象 |
| **极小字号** | 正文字号 < 10.0 pt（难以在投影或会议屏阅读） | High (高危) | 强制提级至 11.5 pt，并同步微调行距 |
| **字号偏小** | 正文字号在 10.0 ~ 11.5 pt 之间 | Low (警告) | 建议优化至 11.5 pt 以上 |
| **文本溢出与截断** | 文本量大但文本框宽高不足，字符密度过高 | High (高危) | 开启 `word_wrap=True`，收紧内边距（上下4pt、左右6pt） |
| **网格微错位抖动** | 并排文本框/卡片垂直 y 轴偏差 1 ~ 12 pt | Medium (中危) | 智能吸附至同一水平基准线（消灭12pt内抖动） |
| **字体族发散** | 全篇出现 4 种以上不同字体族 | Medium (中危) | 统一归并为预设标题族与正文族 |

---

## 3. 多轮自主闭环操作规范 (Self-Calling Loop)

在 DSH / Codex 环境中，模型必须遵循自主闭环流：

```text
[用户提供 PPTX]
       ↓
【Step 1: 基线审计】调用 scripts/audit_pptx.py --json
       ↓  (输出当前健康分、违规形状列表、错位统计)
【Step 2: 针对性微创精修】调用 scripts/polish_pptx.py
       ↓  (自动修复极小字体、消除网格微错位、加固换行与内边距)
【Step 3: 即时复审对账】调用 scripts/audit_pptx.py 
       ↓  (验证分值是否提升至 95+ 分，确认无残余违规)
【Step 4: 高清走查导出】调用 scripts/render_pptx.py 
       ↓  (通过本地 COM 导出高清 PNG，供视觉确认)
[输出对比台账交付]
```

### 快捷一键多轮自治命令
直接调用 `multi_pass_runner.py` 执行完整的自检自修流程：
```bash
python scripts/multi_pass_runner.py input.pptx -o output_polished.pptx --theme swiss --render-dir ./renders --max-rounds 3
```
