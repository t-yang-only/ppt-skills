# PPT-Skills 🎨📊
> **Next-Gen AI Presentation Design, In-Place Refactoring & Autonomous Polish Engine**  
> 全功能旗舰演示文稿生产与存量美化重构引擎：聚合顶流开源 PPT 设计规范、40+ 矢量图表库、工程数据扫描萃取、网络配图审美质检、异形图片蒙版裁切、局部智能抠图去白边、完美生图提示词工程与多维可视化编辑工作台。

[![Python 3.8+](https://img.shields.io/badge/python-3.8+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Agent Skills Compatible](https://img.shields.io/badge/Agent%20Skills-Ready-green.svg)](https://agentskills.io/)
[![100% Native Editable](https://img.shields.io/badge/Objects-100%25%20Editable%20PPTX-brightgreen.svg)]()
[![Visual Canvas Ready](https://img.shields.io/badge/Visual%20Editor-Integrated-blueviolet.svg)]()

---

## 🌟 为什么选择 PPT-Skills？ (Why PPT-Skills?)

传统 AI 生成 PPT 普遍面临四大硬伤：
1. **生成即不可控**：输出大量整页扁平截图或生硬模板套皮，文字截断、元素错位、内容失真。
2. **存量无法拯救**：用户手上已有半成品或汇报文案，AI 无法在保留其完整逻辑和原生对象的前提下进行精准、安全的审美重构。
3. **数据需手动搬运**：工程技术事实、性能数据、对比表格无法直接从代码库和文档中提炼转化为精美排版。
4. **配图死板方框与劣质脏底**：配图只能插入呆板矩形，且网络图片模糊、带水印或发灰；生图素材背景杂乱白边严重，破坏幻灯片美感。

**PPT-Skills** 专为解决上述问题而生。它不仅支持从零到一的交付级工业化制作（Brief-to-PNG），更拥有强大的**已有工程数据扫描萃取、网络配图审美多维质检、存量 PPT 逐页深度审计、异形图片蒙版裁切、局部智能抠图去白边、基于参考图的完美提示词生图、微创美化重塑与多维可视化编辑闭环**能力。

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
│   ├── project_data_harvester.py # 项目代码与文档深度扫描、KPI/表格事实萃取
│   ├── web_image_curator.py# 网络高清配图智能检索、拉普拉斯清晰度与对比度硬性质检
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
    │   ├── taskpkg-data-mining-and-web-curation.md # 项目数据挖掘与网络配图审美质检规范
    │   ├── taskpkg-ppt-design-skill.md     # PPT-Design-Skill 构建与渲染规范
    │   ├── taskpkg-ppt-master.md           # 18 套审美主题与 40+ SVG 图表规范
    │   ├── taskpkg-cyberppt.md             # 咨询级 SCR 论证与数据排版
    │   ├── taskpkg-deck-scoring-benchmark.md # 真实语料客观量化评分基准
    │   ├── taskpkg-irregular-shapes-and-matting.md # 异形图片蒙版、不规则抠图与消白边规范
    │   ├── taskpkg-vision-prompt-and-visual-editing.md # 视觉生图提示词工程与多维可视化编辑规范
    │   ├── taskpkg-organic-design.md       # 有机呼吸感去模板化设计
    │   └── taskpkg-slides-polish.md        # 逐页微创精修与质检规范
    └── qa-checklist.md     # 交付前全要素硬性验收清单
```

---

## 🚀 八大核心工作流模式 (Workflow Modes)

### 1. 项目数据扫描与结构化事实生成 (`harvest-mode`)
- 自动递归扫描工程目录与项目文档。
- 提炼代码行数、核心技术栈、4 组大字报 Hero KPI 指标卡。
- 提取“传统痛点 vs 升级方案”多维对比矩阵，并一键生成原生美化表格入页。
- 自动提取研发演进时间线与版本 Milestone。

### 2. 网络配图检索与审美多维质检 (`web-curate-mode`)
- 关键词检索商业级高清配图素材。
- **四道硬性门禁**：分辨率（≥ 800x500）、长宽比失衡（≤ 3.2:1）、OpenCV 拉普拉斯方差清晰度（$\ge 70.0$）、RMS 对比度（$\ge 28.0$）。
- 确保只有高清、高对比度、无明显水印和失焦的优质图才能装配入幻灯片。

### 3. 异形蒙版与不规则素材装配 (`matting-mask-mode`)
- **9 种非标高审美蒙版**：有机流体（Blob）、苹果超椭圆（Squircle）、六边形、八边形、菱形、水滴、杂志斜切卡片、自由多边形，带平滑抗锯齿边界与描边。
- **局部截取与智能去背精修**：ROI 框选、不规则多边形套索、GrabCut 自动去背，并执行 1~2px 形态学收缩去白边（Defringing）。

### 4. 交付级全流程制作 (`build-mode` - 基于 PPT-Design-Skill)
- 严格遵循 `brief.md` -> `acceptance-contract.md` -> `page-plan.md` -> `visual-direction.md` -> `build.py` 工业级交付链。
- 引入 `MUST` / `SHOULD` / `NICE_TO_HAVE` 三级视觉验收契约，支持 VI 品牌锁定。
- 最终必须完成 `PPTX -> PDF -> PNG` 真实视网膜审查。

### 5. 存量全量美化重构 (`polish-mode` - 基于 ppt-master)
- 统一字体族（严控全篇 ≤ 2 种字体）。
- 标准化字号阶梯（大标题 28~36pt，卡片标题 18~22pt，正文 13~15pt，**严禁正文 < 10.0pt**）。
- 12 列网格卡片化容器规整，杜绝漂移文本框。
- 一键注入 18 套高级审美主题系统。

### 6. 基于参考图的完美提示词生图 (`prompt-craft-mode`)
- **7 维视觉解构**：核心主体、渲染媒介 (3D玻璃拟态/极简黏土/工业金属/扁平杂志)、材质物理、光路、轴测视角与色板提取。
- **专为透明抠图定制**：一键生成纯白无干扰影棚隔离背景提示词，生成的素材专供无损抠出透明背景。

### 7. 多维可视化检视与编辑 (`visual-edit-mode`)
- **本地 Web 交互画板 (`visual_editor.py`)**：12 列网格标尺、5% 安全边距参考线、平移与无级缩放。
- **平台级 Univer Slide**：通过 DSH `sidebar_open` 在侧边栏直接可视化交互编辑矢量幻灯片。
- **桌面级 GenOffice**：调用 `genoffice open` 调起本地原生可视化编辑器。

### 8. 客观量化审计打分与多轮自主闭环 (`audit-mode` & `multi-pass-mode`)
- 扫描整篇 PPT，输出 0~100 分量化健康分与高危缺陷清单。
- 自动调度【基线审计 -> 针对性重构 -> 即时复查 -> 微创修偏 -> COM 高清渲染】，直至评分突破 95 分。

---

## 🛠️ 快速上手与 CLI 指南 (Quickstart)

### 环境依赖
```bash
pip install -r requirements.txt
```

### 1. 扫描已有工程并生成结构化 PPT 数据
```bash
# 扫描工程并直接将结构化对比表注入 PPTX 第 3 页
python scripts/project_data_harvester.py "D:/project/my-repo" --pptx deck.pptx --slide 3
```

### 2. 搜索网络配图、审美质检合格后直接插入 PPT
```bash
# 检索云计算架构配图，经拉普拉斯方差清晰度与对比度质检通过后直接插入
python scripts/web_image_curator.py --query "enterprise cloud architecture" --pptx deck.pptx --slide 2 --shape blob
```

### 3. 异形图片蒙版裁切
```bash
python scripts/shape_image_mask.py hero.png --shape blob --border 2 --border-color "#0066FF" --pptx deck.pptx --slide 1
```

### 4. 局部截取、智能抠图去白边
```bash
python scripts/matting_cutout.py raw.png -o cutout.png --crop "0.2,0.1,0.8,0.9" --auto-bg --pptx deck.pptx --slide 2
```

### 5. 启动本地可视化画板工作台
```bash
python scripts/visual_editor.py deck.pptx --port 8765
```

### 6. 一键多轮自主闭环精修
```bash
python scripts/multi_pass_runner.py input.pptx -o output_perfect.pptx --theme swiss --render-dir ./renders
```

---

## 📜 许可证 (License)

本项目采用 [MIT License](LICENSE) 开源。内部整合的开源子系统遵循其各自原始开源协议。
