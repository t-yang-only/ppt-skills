# taskpkg-ppt-design-skill: 交付级演示文稿架构与渲染引擎规范

> 整合开源顶级规范 `sunchaokun/PPT-Design-Skill`。
> 核心定位：提供从 Brief 需求确认、验收契约（Acceptance Contract）、视觉定向（Visual Direction）、排版配方（Composition Recipes）、渲染包（Rendering Packs）到 PPTX->PDF->PNG 像素级视觉审查的完整交付级设计工作流。

---

## 1. 核心理念与设计原则

1. **受众与目标先行 (Audience First)**：在决定布局前，先确定受众必须理解、记住、感受或执行的内容。
2. **克制优于装饰 (Restraint Over Decoration)**：一套清晰统一的强调色系统远胜于散乱杂糅的视觉特效；所有元素必须有存在理由。
3. **系统性思维 (Systematic Thinking)**：整个演示文稿是一套统一的视觉系统，页边距、节奏韵律、字体体系、色彩角色、圆角与图像处理全局锁死。
4. **意图明确的版式变化 (Intentional Variation)**：仅当故事节奏变化时调整页面架构与信息密度，而非为了“看起来不一样”而随机变化。
5. **严禁仅凭代码运行成功冒充完成**：必须完成 `PPTX -> PDF -> PNG` 渲染走查，由模型直接检验最终 PNG 渲染视觉效果。

---

## 2. 三种核心生成模式 (Generation Modes)

1. **FreeStyle 模式 (`generate_ppt()`)**：
   - 适用于快速概念原型、探索性草稿或简短陈述。
   - 输入结构化 content 数据，快速生成基础版式。
2. **Build Mode (工业级构建模式)**：
   - 严格遵循 `brief.md` -> `acceptance-contract.md` -> `page-plan.md` -> `visual-direction.md` -> `build.py` 交付链。
   - 使用精细参数控制卡片网格、图表容器、字号阶梯与间距比例。
3. **VI Build Mode (品牌视觉识别锁定模式)**：
   - 加载 `theme-lock.yaml` 严格锁定品牌资产（Logo、标准字体、企业主色与辅助色）。
   - 具备品牌一致性审计守卫，禁止越权覆盖品牌视觉规范。

---

## 3. 视觉渲染包与配方库 (Rendering Packs & Composition Recipes)

在 `tools/ppt-design-skill/skill/references/` 中内置多套经实战验证的渲染系统：

### 视觉渲染包 (Visual Rendering Packs)
- `dark-cinematic-tech.json`：暗色电影级科技质感，适配前沿 AI、数据基座与云计算架构。
- `editorial-infrastructure-ledger.json`：杂志编辑风格账本排版，适配宏观经济、供应链、企业战略。
- `scientific-atlas.json`：科研图谱与临床证据级严谨排版，适配生物医药、学术论文、实验图表。

### 排版配方 (Composition Recipes)
- `c1-01`：单核心英雄大指标 + 三分支细分卡片。
- `c2-01`：左右非对称双栏对撞，左叙事/右事实图表。
- `c3-01`：三阶段递进式时间轴与能力里程碑。
- `c4-01`：四象限/矩阵对比与多维度能力评估。

---

## 4. 工具脚本与运行指令

位于 `tools/ppt-design-skill/`:
- `scripts/inspect_pptx.py`：PPTX 结构、文本框、形状与元素层次解析器。
- `scripts/audit_pptx.py`：版式合规与设计缺陷检测。
- `scripts/render_pptx.ps1`：PPTX 转 PDF/PNG 自动化脚本。
- `installer/install.py`：跨平台技能依赖环境检测与安装。
