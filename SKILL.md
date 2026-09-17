---
name: ppt-skill
description: 全功能开源旗舰演示文稿生产与存量美化重构引擎（深度整合 ppt-design-skill, ppt-master, cyber-ppt, deck-scoring-benchmark、项目数据扫描萃取、网络配图审美质检、异形蒙版裁剪、不规则智能抠图、完美生图提示词工程与多维可视化编辑体系）。
metadata:
  short-description: 交付级PPT架构设计、存量美化重构、项目数据萃取、异形抠图装配、客观量化审计与全流程自主闭环
---

# PPT Skill (开源高星全功能升级版)

本 Skill 是面向 **从零到一高标准演示文稿生成、存量已有 PPTX 美化重构、已有工程数据扫描萃取、网络配图审美多维质检、异形图片蒙版裁切、局部智能抠图装配、基于参考图的完美提示词生图、客观量化审美审计与多维可视化编辑闭环** 的全能演示文稿工程体系。

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
    ├── README.md           # 项目主说明文档
    ├── requirements.txt    # 核心依赖清单
    ├── scripts/            # 核心自动化执行与自治闭环脚本
    │   ├── audit_pptx.py       # 逐页客观量化健康审计与违规检测
    │   ├── polish_pptx.py      # 18 套高级审美主题重塑与微创精修
    │   ├── organic_polish.py   # 内容驱动·去模板化·有机呼吸感排版进化引擎
    │   ├── project_data_harvester.py # 项目代码与文档深度扫描、KPI/表格事实萃取
    │   ├── web_image_curator.py# 网络高清配图智能检索、拉普拉斯清晰度与对比度硬性质检
    │   ├── shape_image_mask.py # 异形图片蒙版裁切（有机流体Blob、苹果超椭圆、正六边形、斜切卡片等）
    │   ├── matting_cutout.py   # 局部截取、不规则套索抠图、去背景精修与白边消除
    │   ├── image_prompt_craft.py # 图像特征解构与完美生图提示词工程（专为PPT透明素材定制纯白底提示词）
    │   ├── visual_editor.py    # 本地轻量化 Web 可视化检视工作台（12列网格、安全边距框、缩放平移）
    │   ├── render_pptx.py      # 本地 PowerPoint COM 引擎 1080p 原生无损渲染
    │   └── multi_pass_runner.py# 多轮自主闭环执行器（审计->美化->复审->终验）
    ├── tools/              # 归属母Skill下的专业高星开源子系统
    │   ├── ppt-design-skill/   # 顶级工业级构建工作流（Brief->验收契约->Build Mode->PNG走查）
    │   ├── ppt-master/         # 40+ 矢量 SVG 图表库、14 部经典设计理论、Keynote 案例拆解
    │   ├── cyber-ppt/          # 麦肯锡/BCG 咨询级高密度 SCR 论证框架与严谨排版
    │   ├── deck-scoring-benchmark/ # 真实语料量化评分基准（40分硬指标 + 60分专家评审 Rubric）
    │   ├── ui-ux-pro-max/      # 顶级 UI/UX 规范、192 套专业色卡、74 套字体排版标尺
    │   ├── genoffice/          # Office 命令行引擎（PPTX/DOCX/PDF 跨格式操作与本地编辑）
    │   ├── bggg-creator-image2ppt/ # 逆向还原：图片/截图一键拆解并重建为原生可编辑对象
    │   ├── apikey-image-gen/   # 视觉资产：AI 高清示意图、概念配图生成
    │   ├── pdf/                # 矢量比对：Poppler 渲染走查与高保真图文提取
    │   └── screenshot/         # 屏幕捕获与视觉走查证据收集
    └── references/         # 规范手册、设计理论与验收清单
        ├── subskills/          # 各子系统专业规范说明
        │   ├── taskpkg-data-mining-and-web-curation.md # 项目数据挖掘与网络配图审美质检规范
        │   ├── taskpkg-ppt-design-skill.md     # PPT-Design-Skill 架构与构建模式
        │   ├── taskpkg-ppt-master.md           # 18 套审美主题与 40+ SVG 图表规范
        │   ├── taskpkg-cyberppt.md             # 咨询级 SCR 论证与数据排版
        │   ├── taskpkg-deck-scoring-benchmark.md # 真实语料打分基准与门禁
        │   ├── taskpkg-irregular-shapes-and-matting.md # 异形图片蒙版、不规则抠图与消白边规范
        │   ├── taskpkg-vision-prompt-and-visual-editing.md # 视觉生图提示词工程与多维可视化编辑规范
        │   ├── taskpkg-organic-design.md       # 有机去模板化呼吸排版设计
        │   └── taskpkg-slides-polish.md        # 逐页微创精修与反模式修复
        └── qa-checklist.md     # 交付前全要素硬性验收清单
```

---

## 核心工作流模式 (Workflow Modes)

根据输入与任务目标，自由选择或组合以下核心工作模式：

### 1. 项目数据扫描与结构化事实生成 (`harvest-mode`)
- **适用场景**：已有成熟软件工程、业务文档或科研项目，需要快速将其技术事实转化为高质量演示内容。
- **自动挖掘维度**：
  - 代码行数、核心文件结构、主要依赖与技术栈总结；
  - 核心性能指标（QPS、延迟、可用性、准确率等）转化为大字报 Hero KPI 卡片；
  - 传统痛点 vs 升级方案结构化对比矩阵，一键生成为原生美化表格入页；
  - 研发时间线与版本演进 Milestone 自动提取。

### 2. 网络配图检索与审美多维质检 (`web-curate-mode`)
- **适用场景**：根据 PPT 主题在网络上检索高清商业配图，并自动杜绝模糊、水印、畸变废片。
- **四道硬性门禁**：
  - 分辨率门禁（≥ 800x500）；
  - 畸变门禁（长宽比 ≤ 3.2:1）；
  - 清晰度门禁（OpenCV 拉普拉斯方差 $\ge 70.0$，过滤失焦模糊）；
  - 动态对比度门禁（RMS 标准差 $\ge 28.0$，剔除灰蒙废片）。
- 合格后支持一键联动抠图或异形蒙版，直接装配入幻灯片。

### 3. 异形蒙版与不规则素材装配 (`matting-mask-mode`)
- **9 种非标高审美蒙版**：有机流体（Blob）、苹果超椭圆（Squircle）、六边形、八边形、菱形、水滴、杂志斜切卡片、自由多边形，带平滑抗锯齿边界与描边。
- **局部截取与智能去背精修**：ROI 框选、不规则多边形套索、GrabCut 自动去背，并执行 1~2px 形态学收缩去白边（Defringing）。

### 4. 交付级全流程制作 (`build-mode` - 基于 PPT-Design-Skill)
- 严格遵循 `brief.md` -> `acceptance-contract.md` -> `page-plan.md` -> `visual-direction.md` -> `build.py` -> `PPTX -> PDF -> PNG` 工业级视网膜审查。

### 5. 存量全量美化重构 (`polish-mode` - 基于 ppt-master)
- 统一字体族（全篇 ≤ 2 种）、标准化字号阶梯、12 列网格卡片容器化、一键注入 18 套高级审美主题系统。

### 6. 多维可视化检视与编辑 (`visual-edit-mode`)
- 本地 Web 交互画板（12列网格、5%安全边距、缩放平移）、DSH Univer Slide 侧边栏交互编辑与 GenOffice 桌面可视化编辑。

### 7. 客观量化审计打分 (`audit-mode` - 基于 deck-scoring-benchmark)
- 扫描整篇 PPT，计算假原生纯图片贴图率、极小字号 (<10pt)、文本截断溢出、并排元素微错位 (1~12pt 抖动) 与色彩纯度，输出 0~100 分量化健康卡。

### 8. 多轮自主闭环 (`multi-pass-mode`)
- 自动调度【基线审计 -> 针对性重构 -> 即时复查 -> 微创修偏 -> COM 高清渲染】，直至评分突破 95 分。

---

## 审美与排版硬性红线 (Hard Invariants)

1. **严禁假原生**：交付物必须是 100% 原生可编辑 PPT 对象。严禁将设计稿或外部截图整页烘焙贴入。
2. **素材纯净与画质过关**：网络配图必须通过拉普拉斯清晰度（$\ge 70$）与对比度（$\ge 28$）硬性质检；插入卡片的素材必须经过去白边精修或异形裁切。
3. **字体纪律**：全篇字体族不得超过 3 种（推荐 1 种无衬线标题族 + 1 种正文族）。
4. **字号红线**：页面主标题 28~36pt、模块小标题 18~22pt、正文一级要点 13~15pt，严禁任何主要内容文字低于 10.0 pt。
5. **零溢出加固**：所有文本框必须开启 `word_wrap = True`，内边距控制在上下 4pt、左右 6pt 以内。
6. **相称走查验证**：任何生成或修改，必须通过本地渲染验证（COM 引擎渲染或 PDF/PNG 走查），核查实际版面质量。
