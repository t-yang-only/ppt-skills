# references 目录索引

> 这是ppt-master技能的文档索引入口。告诉你每个文件在哪里、什么时候该读。

---

## 目录结构

```
references/
├── README.md              ← 你在这里
├── cases/                 ← 案例库（好案例/坏案例）
├── guide/                 ← 核心方法论
├── scenarios/             ← 场景指南
└── [其他文件]             ← 单文件文档
```

---

## 快速导航

| 你需要做什么 | 去哪里 |
|------------|--------|
| 看好的PPT案例 | [cases/](#cases-案例库) |
| 看差的PPT案例 | [cases/](#cases-案例库) |
| 理解设计原则 | [guide/](#guide-核心方法论) |
| 按场景制作PPT | [scenarios/](#scenarios-场景指南) |
| 不知道去哪 | 看这个README |

---

## cases 案例库

> 学习好的，避免差的。

### ⭐ 顶级发布会案例

| 文件 | 一句话 | 适合场景 |
|------|--------|---------|
| [huawei-keynote拆解.md](cases/huawei-keynote拆解.md) | 华为发布会设计分析 | 产品发布/路演 |
| [apple-keynote拆解.md](cases/apple-keynote拆解.md) | 苹果发布会设计分析 | 产品发布/路演 |
| [ted-talk-design.md](cases/ted-talk-design.md) | TED演讲PPT设计原则 | 知识分享/培训 |

### ⭐ 失败案例分析

| 文件 | 一句话 | 适合场景 |
|------|--------|---------|
| [failed-bp-5cases.md](cases/failed-bp-5cases.md) | 5个真实BP失败案例 | 融资/商业计划 |

### ⭐ 好vs差对比

| 文件 | 一句话 | 适合场景 |
|------|--------|---------|
| [10-failure-patterns.md](cases/10-failure-patterns.md) | 10个经典失败模式 | 自查/避免错误 |
| [same-scene-good-vs-bad.md](cases/same-scene-good-vs-bad.md) | 同一场景好vs差对比 | 对照学习 |

### ⭐ 实战案例(从理论到代码的桥梁)

| 文件 | 一句话 | 适合场景 |
|------|--------|---------|
| **[real-project-example/](cases/real-project-example/)** | **A 真实项目代码（10 页产品发布 PPT）**，可 `node compile.js` 直接跑 | 学"好设计怎么落地成代码" |

> 💡 **3 类案例互补**:
> - 理论案例(华为/苹果/TED 拆解)→ 学"什么是好设计"
> - 反面案例(failed-bp / 10-failure-patterns)→ 学"什么是不能做的"
> - 实战案例(本目录)→ 学"好设计怎么变成可运行代码"

---

## guide 核心方法论

> 理解"为什么这样做"比"怎么做"更重要。

| 文件 | 一句话 | 什么时候读 |
|------|--------|-----------|
| [design-principles.md](design-principles.md) | CRAP四原则+认知心理学 | 设计决策时 |
| [color-system.md](color-system.md) | 配色7:2:1法则 | 选择配色时 |
| [layout-system.md](layout-system.md) | 版式选择决策树 | 选择版式时 |
| [typography-guide.md](typography-guide.md) | 字体匹配决策矩阵 | 选择字体时 |
| [universal-principles-of-design.md](universal-principles-of-design.md) | 10条认知心理学设计原则 | 理解原理时 |

---

## scenarios 场景指南

> 不同场景需要不同的设计策略。

| 文件 | 场景 | 一句话 |
|------|------|--------|
| [academic.md](scenarios/academic.md) | 学术汇报 | 信息密度高，逻辑严谨 |
| [commercial.md](scenarios/commercial.md) | 商业路演 | 结论先行，数据说话 |
| [training.md](scenarios/training.md) | 企业培训 | 循序渐进，案例丰富 |
| [workplace.md](scenarios/workplace.md) | 工作汇报 | 结构清晰，重点突出 |

---

## 其他文档

| 文件 | 一句话 |
|------|--------|
| [case-studies.md](case-studies.md) | 案例索引入口（→ 指向cases/） |
| [charts.md](charts.md) | 图表设计指南 |
| [common-errors.md](common-errors.md) | 常见错误清单 |
| [editing.md](editing.md) | PPT编辑技巧 |
| [feedback-loop.md](feedback-loop.md) | 演示反馈收集与改进 |
| [master-inspiration.md](master-inspiration.md) | 大师灵感50例 |
| [philosophy.md](philosophy.md) | PPT设计哲学 |
| [pptxgenjs.md](pptxgenjs.md) | PptxGenJS API参考 |
| [quality-checklist.md](quality-checklist.md) | 质量检查清单 |
| [scenarios.md](scenarios.md) | 场景总览 |
| [slide-types.md](slide-types.md) | 幻灯片类型指南 |
| [speech-preparation.md](speech-preparation.md) | 演讲准备指南 |
| **[svg-generation-guide.md](svg-generation-guide.md)** | **SVG 生成实战指南**（代码示例、技术规范） |
| **[svg-graphics-guide.md](svg-graphics-guide.md)** | **52种 SVG 模板说明**（分类、场景、决策树） |
| [techniques.md](techniques.md) | 技巧集合 |
| [thinking.md](thinking.md) | 设计思维 |
| [three-dimensions.md](three-dimensions.md) | PPT三维体系 |
| [narrative-arc.md](narrative-arc.md) | **叙事弧模板**（内容结构化方法论） |
| [accessibility.md](accessibility.md) | 无障碍设计检查清单 |
| [extended-themes.md](extended-themes.md) | 扩展配色方案参考 |
| **[slide-techniques.md](slide-techniques.md)** | **实战幻灯片技巧**（slideConfig/假渐变/数据驱动等,来自真实项目沉淀） |

---

## 按问题查找

| 问题 | 推荐文件 |
|------|---------|
| 不知道什么是"好PPT" | cases/huawei-keynote拆解.md |
| 总被说"内容太多" | cases/ted-talk-design.md |
| 不知道投资人想看什么 | cases/failed-bp-5cases.md |
| 不知道用什么配色 | guide/color-system.md |
| 不知道用什么字体 | guide/typography-guide.md |
| 不知道用什么版式 | guide/layout-system.md |
| 逻辑不清晰 | guide/design-principles.md |
| 容易被忽略的错误 | cases/10-failure-patterns.md |
| 不知道这个场景该怎么设计 | scenarios/[对应场景].md |
| 需要画复杂图形/图表 | svg-generation-guide.md |
| 不知道选什么 SVG 模板 | svg-graphics-guide.md |
| **不知道内容怎么组织** | **narrative-arc.md** |
| **写 slide-XX.js 时不知道代码模式** | **slide-techniques.md** |
| **想看真实项目代码长什么样** | **cases/real-project-example/** |

---

## SVG 图形模板库

> 52个SVG图形模板，位于技能根目录的 `templates/svg-charts/`

**快速查找**：
| 用途 | 推荐模板 |
|------|----------|
| 排名对比 | horizontal_bar_chart, bar_chart, pareto_chart |
| 趋势变化 | line_chart, area_chart |
| 占比组成 | donut_chart, pie_chart |
| KPI展示 | kpi_cards, bullet_chart |
| 转化漏斗 | funnel_chart, sankey_chart |
| 战略框架 | swot_analysis, porter_five_forces, matrix_2x2 |
| 信息图示 | pyramid_chart, venn_diagram, mind_map, word_cloud |

**详细索引**：`templates/svg-charts/charts_index.json`

---

## 更新日志

| 日期 | 更新内容 |
|------|----------|
| 2026-06-01 | 新增实战案例目录 `cases/real-project-example/`（A 真实 10 页产品发布项目代码，可 `node compile.js` 直接跑） |
| 2026-06-01 | 新增实战幻灯片技巧（slide-techniques.md），从 A 的 work/minimax-code 项目沉淀 |
| 2026-04-29 | 新增叙事弧模板（narrative-arc.md） |
| 2026-04-29 | 新增无障碍设计清单（accessibility.md） |
| 2026-04-29 | 新增扩展配色方案（extended-themes.md） |
| 2026-04-18 | **修复关键缺陷**：新增安全边距规范和网格系统（pptxgenjs.md），新增CRAP原则代码检查清单（SKILL.md 4.4节） |
| 2026-04-18 | 新增 SVG 高级图形文档索引（svg-generation-guide.md、svg-graphics-guide.md） |
| 2026-03-30 | 新增cases/目录，扁平化案例结构 |
| 2026-03-30 | 新增统一索引入口README.md |
| 2026-03-30 | 整理为cases/guide/scenarios三分类 |
