---
name: ppt-skill
description: 全功能开源旗舰演示文稿生产与存量美化重构引擎（深度整合 ppt-design-skill, ppt-master, cyber-ppt, deck-scoring-benchmark 等开源高星工具与理论体系）。
metadata:
  short-description: 交付级PPT架构设计、存量美化重构、客观量化审计与全流程自主闭环
---

# PPT Skill (开源高星全功能升级版)

本 Skill 是面向 **从零到一高标准演示文稿生成、存量已有 PPTX 美化重构、客观量化审美审计与多轮自主闭环** 的全能演示文稿工程体系。

深度整合 GitHub 开源顶流 PPT 体系（`sunchaokun/PPT-Design-Skill`、`Categorytyy/ppt-master`、`cyber-ppt`、`deck-scoring-benchmark` 等），为专业汇报、商业路演、学术答辩及高难度展示提供系统级工程支撑与有机演进空间。

---

## 核心能力架构 (Unified Architecture)

母 Skill `ppt-skills` (即 `ppt-skill`) 统一调度内部专业工具集与算法脚本：

```text
D:\skills\
├── auto-skills/            # [母Skill] 自动化与工作流调度总控
├── hacker-skills/          # [母Skill] 逆向与安全分析工具包
├── nm-skills/              # [母Skill] 纳米自动化与工程协作台账
└── ppt-skills/ (ppt-skill) # [母Skill] 演示文稿生产与存量美化总控中枢
    ├── SKILL.md            # 本文件：全局调度规范与操作指南
    ├── scripts/            # 核心自动化执行与自治闭环脚本
    │   ├── audit_pptx.py       # 逐页客观量化健康审计与违规检测
    │   ├── polish_pptx.py      # 18 套高级审美主题重塑与微创精修
    │   ├── organic_polish.py   # 内容驱动·去模板化·有机呼吸感排版进化引擎
    │   ├── render_pptx.py      # 本地 PowerPoint COM 引擎 1080p 原生无损渲染
    │   └── multi_pass_runner.py# 多轮自主闭环执行器（审计->美化->复审->终验）
    ├── tools/              # 归属母Skill下的专业高星开源子系统
    │   ├── ppt-design-skill/   # 顶级工业级构建工作流（Brief->验收契约->Build Mode->PNG走查）
    │   ├── ppt-master/         # 40+ 矢量 SVG 图表库、14 部经典设计书理论、顶级 Keynote 拆解
    │   ├── cyber-ppt/          # 麦肯锡/BCG 咨询级高密度 SCR 论证框架与严谨排版
    │   ├── deck-scoring-benchmark/ # 真实语料量化评分基准（自动指标 + 专家评审 Rubric）
    │   ├── ui-ux-pro-max/      # 顶级 UI/UX 规范、192 套专业色卡、74 套字体排版标尺
    │   ├── genoffice/          # Office 命令行引擎（PPTX/DOCX/PDF 跨格式操作）
    │   ├── bggg-creator-image2ppt/ # 逆向还原：图片/截图一键拆解并重建为原生可编辑对象
    │   ├── apikey-image-gen/   # 视觉资产：AI 高清示意图、概念配图生成
    │   ├── pdf/                # 矢量比对：Poppler 渲染走查与高保真图文提取
    │   └── screenshot/         # 屏幕捕获与视觉走查证据收集
    └── references/         # 规范手册、设计理论与验收清单
        ├── subskills/          # 各子系统专业规范说明
        │   ├── taskpkg-ppt-design-skill.md     # PPT-Design-Skill 架构与构建模式
        │   ├── taskpkg-ppt-master.md           # 18 套审美主题与 40+ SVG 图表规范
        │   ├── taskpkg-cyberppt.md             # 咨询级 SCR 论证与数据排版
        │   ├── taskpkg-deck-scoring-benchmark.md # 真实语料打分基准与门禁
        │   ├── taskpkg-organic-design.md       # 有机去模板化呼吸排版设计
        │   └── taskpkg-slides-polish.md        # 逐页微创精修与反模式修复
        └── qa-checklist.md     # 交付前全要素硬性验收清单
```

---

## 核心工作流模式 (Workflow Modes)

根据输入与任务目标，自由选择或组合以下五大工作模式：

### 1. 交付级全流程制作 (`build-mode` - 基于 PPT-Design-Skill)
- **适用场景**：高要求路演、大会发布、重要竞标、学术前沿展示。
- **流程链条**：
  1. 确认需求与大纲 (`brief.md`)。
  2. 建立**验收契约** (`acceptance-contract.md`)，明确 MUST / SHOULD / NICE_TO_HAVE 视觉门禁。
  3. 确定视觉方向与渲染包 (`visual-direction.md` / `Rendering Packs`)。
  4. 编写构建脚本 (`build.py`) 并生成原生可编辑 PPTX。
  5. 自动化导出 `PPTX -> PDF -> PNG`，进行直接视觉审查。

### 2. 存量全量美化重构 (`polish-mode` - 基于 ppt-master)
- **适用场景**：已有草稿 PPT、内容完整但排版凌乱、字体发散、审美陈旧。
- **重构动作**：
  - **字体系统归一**：全篇锁定 1 套标题族 + 1 套正文族。
  - **字号层级阶梯**：主标题 28~36pt、卡片标题 18~22pt、正文 13~15pt，严禁正文 < 10.0pt。
  - **网格容器化**：消除零散漂移文本框，重塑 12 列网格卡片容器。
  - **18 套高级审美主题**：一键注入 `swiss`（瑞士极简）、`finance`（商业财经）、`dark-saas`（深色SaaS）、`academic`（学术科研）、`mckinsey`（麦肯锡咨询）、`ink`（新中式水墨）、`competition`（答辩硬核）等。

### 3. 客观量化审计打分 (`audit-mode` - 基于 deck-scoring-benchmark)
- **适用场景**：交付前质检、PPT 现状体检、竞赛评分预估。
- **客观指标计算**：
  - 纯图片整页贴图率（严惩假原生，覆盖率 > 80% 判硬伤）。
  - 超小字号检测（< 10pt 阻断）。
  - 文本框文字截断溢出检测。
  - 并排元素微错位抖动分析（1~12pt 偏差）。
  - XML 遍历 `a:srgbClr` 统计全篇主辅强调色，输出健康评分卡（0~100 分）。

### 4. 有机去模板化排版 (`organic-mode`)
- **适用场景**：告别千篇一律的卡片网格，需要视觉呼吸感与大片杂志质感。
- **有机进化机制**：
  - **大指标强聚焦**：自动识别核心数据指标（增长率、倍数、核心数字），放大至 30pt+ 并赋予重点色彩。
  - **空气感呼吸留白**：收紧文本框内边距释放空间，增加模块间呼吸负空间。
  - **非对称张力排版**：大图与精致数据卡片左右呼应，构建动静结合的视觉动线。

### 5. 多轮自主闭环 (`multi-pass-mode`)
- **适用场景**：无需人工干预的自动化精修闭环。
- **自主循环**：
  - 运行 `python scripts/multi_pass_runner.py input.pptx -o polished.pptx`
  - 自动调度【基线审计 -> 针对性重构 -> 即时复查 -> 微创修偏 -> COM 高清渲染】，直至指标评分突破 95 分。

---

## 40+ 矢量 SVG 图表组件库 (Component Library)

存放在 `tools/ppt-master/ppt-master/templates/svg-charts/`，支持直接调用或嵌入：

| 类别 | 包含组件 |
|---|---|
| **趋势与数据** | `area_chart`, `line_chart`, `dual_axis_line_chart`, `bar_chart`, `grouped_bar_chart`, `horizontal_bar_chart`, `stacked_bar_chart`, `waterfall_chart`, `pareto_chart` |
| **分布与结构** | `pie_chart`, `donut_chart`, `treemap_chart`, `scatter_chart`, `bubble_chart`, `box_plot_chart`, `heatmap_chart` |
| **流程与演进** | `process_flow`, `chevron_process`, `snake_flow`, `cycle_diagram`, `roadmap_vertical`, `gantt_chart`, `timeline`, `numbered_steps` |
| **架构与关系** | `hub_spoke`, `org_chart`, `concentric_circles`, `mind_map`, `fishbone_diagram`, `sankey_chart`, `venn_diagram` |
| **战略与分析** | `swot_analysis`, `matrix_2x2`, `porter_five_forces`, `pyramid_chart`, `funnel_chart`, `radar_chart`, `gauge_chart`, `pros_cons_chart`, `kpi_cards` |

---

## 审美与排版硬性红线 (Hard Invariants)

1. **严禁假原生**：交付物必须是 100% 原生可编辑 PPT 对象。严禁将设计稿或外部截图整页烘焙贴入。
2. **字体纪律**：全篇字体族不得超过 3 种（推荐 1 种无衬线标题族 + 1 种正文族）。
3. **字号红线**：
   - 页面主标题：28 ~ 36 pt (Bold)
   - 模块/卡片小标题：18 ~ 22 pt (Bold)
   - 正文一级要点：13 ~ 15 pt (Regular)
   - 严禁任何主要内容文字低于 10.0 pt。
4. **零溢出加固 (Zero-Overflow Defense)**：所有文本框必须开启 `word_wrap = True`，内边距控制在上下 4pt、左右 6pt 以内。
5. **对齐平滑**：并排排列的卡片或文本框，垂直 y 轴偏差在 12pt 内的必须吸附至同一水平基准线。
6. **相称走查验证**：任何生成或修改，必须通过本地渲染验证（COM 引擎渲染或 PDF/PNG 走查），核查实际版面质量。
