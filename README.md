# PPT-Skills 🎨📊
> **Next-Gen AI Presentation Design, In-Place Refactoring & Autonomous Polish Engine**  
> 全功能旗舰演示文稿生产与存量美化重构引擎，聚合顶流开源 PPT 设计规范、40+ 矢量图表库与客观量化质检系统。

[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Agent Skills Compatible](https://img.shields.io/badge/Agent%20Skills-Ready-green.svg)](https://agentskills.io/)
[![100% Native Editable](https://img.shields.io/badge/Objects-100%25%20Editable%20PPTX-brightgreen.svg)]()

---

## 🌟 为什么选择 PPT-Skills？ (Why PPT-Skills?)

市面上的大多数 AI 生成 PPT 方案存在两个致命痛点：
1. **生成即不可控**：输出大量整页扁平截图或生硬模板套皮，文字截断、元素错位、内容失真。
2. **存量无法拯救**：用户手上已有做好的半成品或汇报文案，AI 无法在保留其完整逻辑和原生对象的前提下进行精准、安全的审美重构。

**PPT-Skills** 专为解决上述问题而生。它不仅支持从零到一的交付级工业化制作（Brief-to-PNG），更拥有强大的**存量 PPT 逐页深度审计、微创美化重塑与多轮自主闭环**能力。

---

## 🏗️ 体系架构 (Architecture)

PPT-Skills 采用模块化母 Skill 结构，内部工具统一收拢在 `tools/` 目录，杜绝平铺冗余，提供纯粹通用的演进空间：

```text
ppt-skills/
├── SKILL.md                # Agent Skill 入口规范与全局调度指令
├── README.md               # 项目主说明文档
├── requirements.txt        # 核心依赖清单
├── .gitignore              # Git 过滤规则
├── scripts/                # 核心自治运行脚本
│   ├── audit_pptx.py       # 逐页客观量化健康审计（纯图片检测/字号下限/溢出/微错位）
│   ├── polish_pptx.py      # 18 套高级审美主题重构、网格容器化与微创精修
│   ├── organic_polish.py   # 内容驱动·去模板化·有机呼吸感排版进化引擎
│   ├── render_pptx.py      # 本地 PowerPoint COM 引擎 1080p 原生无损渲染
│   └── multi_pass_runner.py# 多轮自主闭环执行器（基线审计->美化->复审->终验）
├── tools/                  # 内部集成的顶流开源专业工具箱
│   ├── ppt-design-skill/   # 工业级构建流（Brief->验收契约->Build Mode->PNG走查）
│   ├── ppt-master/         # 40+ 矢量 SVG 图表库、14 部经典设计理论、Keynote 案例拆解
│   ├── cyber-ppt/          # 麦肯锡/BCG 咨询级高密度 SCR 结构化论证与严谨排版
│   ├── deck-scoring-benchmark/ # 真实语料客观量化打分基准（40分硬指标 + 60分专家评审）
│   ├── ui-ux-pro-max/      # 192 套专业色卡、74 套字体排版标尺与设计系统
│   ├── genoffice/          # Office 命令行引擎（PPTX/DOCX/PDF 跨格式操作）
│   ├── bggg-creator-image2ppt/ # 逆向还原：图片/截图一键拆解并重建为原生可编辑对象
│   ├── apikey-image-gen/   # 资产生成：AI 概念插图、示意图与背景素材生成
│   ├── pdf/                # 矢量比对：Poppler 渲染走查与高保真图文提取
│   └── screenshot/         # 屏幕捕获与视觉走查证据收集
└── references/             # 规范手册、设计理论与验收清单
    ├── subskills/          # 专业子系统详细规范说明
    │   ├── taskpkg-ppt-design-skill.md     # PPT-Design-Skill 构建与渲染规范
    │   ├── taskpkg-ppt-master.md           # 18 套审美主题与全量美化规范
    │   ├── taskpkg-cyberppt.md             # 咨询级 SCR 论证与数据排版
    │   ├── taskpkg-deck-scoring-benchmark.md # 真实语料客观量化评分基准
    │   ├── taskpkg-organic-design.md       # 有机呼吸感去模板化设计
    │   └── taskpkg-slides-polish.md        # 逐页微创精修与质检规范
    └── qa-checklist.md     # 交付前全要素硬性验收清单
```

---

## 🚀 五大工作流模式 (Workflow Modes)

### 1. 工业级交付构建 (`build-mode`)
*基于 `sunchaokun/PPT-Design-Skill`*
- 严格遵循 `brief.md` -> `acceptance-contract.md` -> `page-plan.md` -> `visual-direction.md` -> `build.py` 工业级交付链。
- 引入 `MUST` / `SHOULD` / `NICE_TO_HAVE` 三级视觉验收契约，支持 VI 品牌锁定。
- 最终必须完成 `PPTX -> PDF -> PNG` 真实视网膜审查。

### 2. 存量全量美化重构 (`polish-mode`)
*基于 `Categorytyy/ppt-master`*
- 统一字体族（严控全篇 ≤ 2 种字体）。
- 标准化字号阶梯（大标题 28~36pt，卡片标题 18~22pt，正文 13~15pt，**严禁正文 < 10.0pt**）。
- 12 列网格卡片化容器规整，杜绝漂移文本框。
- 一键注入 18 套高级审美主题系统。

### 3. 客观量化审计打分 (`audit-mode`)
*基于 `deck-scoring-benchmark`*
- 扫描整篇 PPT，输出 0~100 分量化健康分与高危缺陷清单。
- 核心检测维度：
  - **假原生检测**：严查纯图片整页贴图（覆盖率 > 80% 判硬伤）。
  - **极小字号检测**：正文低于 10pt 直接触发扣分告警。
  - **零溢出检测**：计算文本框容积与文字换行，排查文字截断。
  - **微错位检测**：扫描并排卡片 1~12pt 水平/垂直视觉抖动。
  - **色彩纯度**：XML 遍历 `a:srgbClr` 审计主题色与强调色收敛度。

### 4. 内容驱动·有机排版 (`organic-mode`)
- 告别千篇一律的机械卡片堆砌，追求杂志级版面节奏。
- **英雄大指标聚焦**：自动识别核心业务指标（增长率、效率倍数、核心数字），放大至 30pt+ 并赋予强调色。
- **空气感留白**：内边距微调释放空间，拉开模块层级，形成视觉呼吸感。
- **非对称张力**：大图与高密度卡片动静呼应，引导自然阅读视线。

### 5. 多轮自主闭环 (`multi-pass-mode`)
- **零人工介入的自动化精修**：
  - 第 1 轮：基线客观量化审计 (Baseline Audit)。
  - 第 2 轮：针对性审美重构 (Theme & Layout Polish)。
  - 第 3 轮：即时复审与微创修偏 (Re-audit & Fine-tuning)。
  - 终验轮：Windows COM 高清渲染走查图与 Before vs After 量化对比报表。

---

## 🎨 18 套内置高级审美视觉系统 (Themes)

| 风格 ID | 风格名称 | 核心色调 (60-30-10) | 推荐场景 |
|---|---|---|---|
| `swiss` | 瑞士网格极简风 | 白底 (60%) + 浅灰卡片 (30%) + 国际橙 (10%) | 科技发布、设计评审、现代极简演示 |
| `finance` | 财经商业杂志风 | 象牙白 + 蓝灰底 + 典雅金 | 商业路演、财报分析、商业计划书 |
| `dark-saas` | SaaS深色玻璃拟态 | 极致黑 + 深蓝灰卡片 + 霓虹青 | 人工智能、云计算、SaaS 产品发布 |
| `academic` | 论文学术科研风 | 纯白 + 浅灰 + 牛津蓝 / 剑桥红 | 硕博答辩、课题结题、严谨研究报告 |
| `mckinsey` | 麦肯锡咨询报告风 | 经典深蓝 + 结构化 Action Title + 珊瑚红 | 战略咨询、管理汇报、行业深度分析 |
| `ink` | 新中式水墨风 | 宣纸暖白 + 浅墨褐 + 丹砂朱红 | 文化出海、文旅推介、精品文化汇报 |
| `competition` | 国赛答辩硬核风 | 科技深蓝 + 指标卡强化 + 架构图原生化 | 挑战杯、互联网+、创新创业大赛答辩 |
| `editorial-bold` | 现代大字报编辑风 | 米白 + 炭黑 + 荧光黄 | 创意提案、品牌定位、趋势洞察 |
| `eco-carbon` | 能源双碳绿色风 | 极淡自然绿 + 灰绿 + 森林深绿 | 新能源、ESG、环保科技汇报 |
| `medical-clean` | 医疗健康科技风 | 洁净白 + 浅天青 + 医学蓝 | 生物医药、医疗器械、健康管理 |

---

## 📊 40+ 矢量 SVG 图表组件库 (Vector Library)

存放在 `tools/ppt-master/ppt-master/templates/svg-charts/`，无需外部依赖即可直接调用嵌入：

- **趋势与数据**：面积图 (`area_chart`)、折线图 (`line_chart`)、双轴折线图 (`dual_axis_line_chart`)、柱状图 (`bar_chart`)、分组柱状图 (`grouped_bar_chart`)、条形图 (`horizontal_bar_chart`)、堆叠柱状图 (`stacked_bar_chart`)、瀑布图 (`waterfall_chart`)、帕累托图 (`pareto_chart`)。
- **分布与结构**：饼图 (`pie_chart`)、环形图 (`donut_chart`)、树状矩形图 (`treemap_chart`)、散点图 (`scatter_chart`)、气泡图 (`bubble_chart`)、箱线图 (`box_plot_chart`)、热力图 (`heatmap_chart`)。
- **流程与演进**：线性流程图 (`process_flow`)、箭头阶段图 (`chevron_process`)、蛇形流程图 (`snake_flow`)、循环图 (`cycle_diagram`)、垂直路线图 (`roadmap_vertical`)、甘特图 (`gantt_chart`)、时间轴 (`timeline`)、步骤标号图 (`numbered_steps`)。
- **架构与关系**：中心辐射图 (`hub_spoke`)、组织架构图 (`org_chart`)、同心圆关系图 (`concentric_circles`)、思维导图 (`mind_map`)、鱼骨图 (`fishbone_diagram`)、桑基图 (`sankey_chart`)、韦恩图 (`venn_diagram`)。
- **战略与分析**：SWOT 分析图 (`swot_analysis`)、波士顿/2x2 矩阵 (`matrix_2x2`)、波特五力模型 (`porter_five_forces`)、金字塔层级图 (`pyramid_chart`)、漏斗图 (`funnel_chart`)、雷达图 (`radar_chart`)、仪表盘 (`gauge_chart`)、优劣势对比图 (`pros_cons_chart`)、KPI 指标卡组 (`kpi_cards`)。

---

## 🛠️ 快速上手与 CLI 指南 (Quickstart)

### 环境依赖
```bash
pip install -r requirements.txt
```

### 1. 快速进行量化健康体检
```bash
python scripts/audit_pptx.py input.pptx --json
```

### 2. 存量 PPT 一键美化重排
```bash
python scripts/polish_pptx.py input.pptx -o output_polished.pptx --theme swiss
```

### 3. 内容驱动·有机去模板化排版
```bash
python scripts/organic_polish.py input.pptx -o output_organic.pptx --theme mckinsey
```

### 4. 导出 1080p 高清走查截图 (Windows COM)
```bash
python scripts/render_pptx.py input.pptx -o ./renders --width 1920 --height 1080
```

### 5. 一键多轮自主闭环精修
```bash
python scripts/multi_pass_runner.py input.pptx -o output_perfect.pptx --theme swiss --render-dir ./renders
```

---

## 🤖 作为 Agent Skill 使用 (Agent Integration)

PPT-Skills 原生遵循 [Agent Skills 开放标准](https://agentskills.io/)，可无缝装载至各大 AI 协作平台：

- **Claude Code**: 安装至 `~/.claude/skills/ppt-skill` 或项目 `.claude/skills/`
- **OpenAI Codex**: 安装至 `~/.codex/skills/ppt-skill`
- **Cursor / OpenCode**: 安装至项目 `.agents/skills/ppt-skill`
- **DeepSeek Harness (DSH)**: 放置于项目根目录 `skills/ppt-skills`

AI Agent 读取 `SKILL.md` 后，即可根据用户自然语言指令，自主在 `build-mode`、`polish-mode`、`audit-mode`、`organic-mode` 与 `multi-pass-mode` 之间做出最优决策调度。

---

## 📜 许可证 (License)

本项目采用 [MIT License](LICENSE) 开源。内部整合的开源子系统遵循其各自原始开源协议。
