# PPT-Skills 🎨📊
> **Next-Gen AI Presentation Design, In-Place Refactoring & Autonomous Polish Engine**  
> 全功能旗舰演示文稿生产与存量美化重构引擎：聚合顶流开源 PPT 设计规范、40+ 矢量图表库、异形图片蒙版裁切、局部智能抠图去白边、完美生图提示词工程与多维可视化编辑工作台。

[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Agent Skills Compatible](https://img.shields.io/badge/Agent%20Skills-Ready-green.svg)](https://agentskills.io/)
[![100% Native Editable](https://img.shields.io/badge/Objects-100%25%20Editable%20PPTX-brightgreen.svg)]()
[![Visual Canvas Ready](https://img.shields.io/badge/Visual%20Editor-Integrated-blueviolet.svg)]()

---

## 🌟 为什么选择 PPT-Skills？ (Why PPT-Skills?)

市面上的大多数 AI 生成 PPT 方案存在三大致命痛点：
1. **生成即不可控**：输出大量整页扁平截图或生硬模板套皮，文字截断、元素错位、内容失真。
2. **存量无法拯救**：用户手上有半成品或汇报文案，AI 无法在保留其完整逻辑和原生对象的前提下进行精准、安全的审美重构。
3. **配图死板方框与脏底**：配图只能插入矩形框，缺乏杂志大片的呼吸感；生成的素材带有杂乱底噪白边，无法完美融入幻灯片卡片。

**PPT-Skills** 专为解决上述问题而生。它不仅支持从零到一的交付级工业化制作（Brief-to-PNG），更拥有强大的**存量 PPT 逐页深度审计、异形图片蒙版裁切、局部智能抠图去白边、基于参考图的完美提示词生图、微创美化重塑与多维可视化编辑闭环**能力。

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
│   ├── shape_image_mask.py # 异形图片蒙版裁切（有机流体Blob、苹果超椭圆、正六边形、斜切卡片等）
│   ├── matting_cutout.py   # 局部截取、不规则套索抠图、去背景精修与白边消除
│   ├── image_prompt_craft.py # 图像特征解构与完美生图提示词工程（专为PPT透明素材定制纯白底提示词）
│   ├── visual_editor.py    # 本地轻量化 Web 可视化检视工作台（12列网格、安全边距框、缩放平移）
│   ├── render_pptx.py      # 本地 PowerPoint COM 引擎 1080p 原生无损渲染
│   └── multi_pass_runner.py# 多轮自主闭环执行器（基线审计->美化->复审->终验）
├── tools/                  # 内部集成的顶流开源专业工具箱
│   ├── ppt-design-skill/   # 工业级构建流（Brief->验收契约->Build Mode->PNG走查）
│   ├── ppt-master/         # 40+ 矢量 SVG 图表库、14 部经典设计理论、Keynote 案例拆解
│   ├── cyber-ppt/          # 麦肯锡/BCG 咨询级高密度 SCR 结构化论证与严谨排版
│   ├── deck-scoring-benchmark/ # 真实语料客观量化打分基准（40分硬指标 + 60分专家评审）
│   ├── ui-ux-pro-max/      # 192 套专业色卡、74 套字体排版标尺与设计系统
│   ├── genoffice/          # Office 命令行引擎（PPTX/DOCX/PDF 跨格式操作与本地编辑）
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
    │   ├── taskpkg-irregular-shapes-and-matting.md # 异形图片蒙版、不规则抠图与消白边规范
    │   ├── taskpkg-vision-prompt-and-visual-editing.md # 视觉生图提示词工程与多维可视化编辑规范
    │   ├── taskpkg-organic-design.md       # 有机呼吸感去模板化设计
    │   └── taskpkg-slides-polish.md        # 逐页微创精修与质检规范
    └── qa-checklist.md     # 交付前全要素硬性验收清单
```

---

## 🚀 七大核心工作流模式 (Workflow Modes)

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

### 3. 异形蒙版与不规则素材装配 (`matting-mask-mode`)
- **告别千篇一律的方形图片框**，提供 9 种不规则高审美蒙版：
  - `blob`：平滑多阶贝塞尔有机流体水滴形，大片级留白破界视觉。
  - `squircle`：苹果级超椭圆曲率 $(x/a)^4 + (y/b)^4 = 1$。
  - `hexagon` / `octagon`：蜂巢正六边形 / 现代八边形，工业与科技严谨感。
  - `slanted_card`：杂志对角斜切卡片，动感倾斜视线。
  - `polygon`：支持任意自由多边形顶点次像素抗锯齿裁切。
- **局部截取与智能抠图精修**：
  - 自由套索框选核心区域，本地 GrabCut 迭代剔除背景。
  - **边缘去白边 (Defringing)**：1~2px 形态学收缩与高斯边缘羽化，消除暗色背景下的白色杂边与光晕。

### 4. 基于参考图的完美提示词生图 (`prompt-craft-mode`)
- **7 维视觉解构**：核心主体、渲染媒介 (3D玻璃拟态/极简黏土/工业金属/扁平杂志)、材质物理 (次表面散射/反射)、光路、轴测视角与色板提取。
- **专为透明抠图定制**：一键生成纯白无干扰影棚隔离背景提示词，确保生成的素材可被 100% 完美抠出透明背景。

### 5. 多维可视化检视与编辑 (`visual-edit-mode`)
- **本地 Web 交互画板 (`visual_editor.py`)**：12 列网格标尺、5% 安全边距参考线、平移与无级缩放。
- **平台级 Univer Slide**：通过 DSH `sidebar_open` 在侧边栏直接可视化交互编辑矢量幻灯片。
- **桌面级 GenOffice**：调用 `genoffice open` 调起本地原生可视化编辑器。

### 6. 客观量化审计打分 (`audit-mode`)
*基于 `deck-scoring-benchmark`*
- 扫描整篇 PPT，输出 0~100 分量化健康分与高危缺陷清单。
- 检测维度：假原生纯图片贴图率、极小字号 (<10pt)、文本截断溢出、并排元素微错位 (1~12pt 抖动)、XML 遍历色相纯度。

### 7. 多轮自主闭环 (`multi-pass-mode`)
- 自动调度【基线审计 -> 针对性重构 -> 即时复查 -> 微创修偏 -> COM 高清渲染】，直至评分突破 95 分。

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

### 1. 异形图片蒙版裁切并植入 PPT
```bash
# 裁切为有机流体形状并注入第 1 页
python scripts/shape_image_mask.py hero.png --shape blob --border 2 --border-color "#0066FF" --pptx deck.pptx --slide 1 --left 2.0 --top 1.5 --width 4.0

# 裁切为苹果超椭圆
python scripts/shape_image_mask.py product.png -o squircle_product.png --shape squircle
```

### 2. 局部截取、智能抠图去白边并植入 PPT
```bash
# 截取局部并自动抠图消白边
python scripts/matting_cutout.py raw.png -o cutout.png --crop "0.2,0.1,0.8,0.9" --auto-bg --pptx deck.pptx --slide 2 --left 4.0 --top 2.0 --width 3.5
```

### 3. 基于参考图反推完美生图提示词
```bash
python scripts/image_prompt_craft.py ref.png --subject "智能工业云平台核心大脑" --style 3d-glass --transparent-asset
```

### 4. 启动本地可视化画板检视工作台
```bash
python scripts/visual_editor.py deck.pptx --port 8765
```

### 5. 存量 PPT 一键美化重排
```bash
python scripts/polish_pptx.py input.pptx -o output_polished.pptx --theme swiss
```

### 6. 一键多轮自主闭环精修
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

AI Agent 读取 `SKILL.md` 后，即可根据自然语言指令自主调度 `build-mode`、`polish-mode`、`matting-mask-mode`、`prompt-craft-mode`、`visual-edit-mode`、`audit-mode` 与 `multi-pass-mode`。

---

## 📜 许可证 (License)

本项目采用 [MIT License](LICENSE) 开源。内部整合的开源子系统遵循其各自原始开源协议。
