# Changelog

所有对 ppt-master 技能的重要变更都会记录在此文件。

本文件格式基于 [Keep a Changelog](https://keepachangelog.com/zh-CN/1.1.0/),
项目版本遵循 [语义化版本](https://semver.org/lang/zh-CN/)。

---

## [Unreleased] - 2026-06-01

### ✨ 新增 (Added)

#### 1. 实战项目案例 `references/cases/real-project-example/`

把"理论案例"和"实战案例"分开,让用户能从完整跑通的项目里学"好设计怎么落地成代码"。

- **完整 10 页产品发布 PPT 代码**:封面/目录/痛点/能力 1-3/场景/对比/总结
- **`compile.js`**:已通用化,4 处可改(theme/title/author/outFile),`node compile.js` 直接跑
- **`README.md`**:说明怎么用、每页能学到什么、跟其他案例的关系
- **junction 依赖**:`node_modules → ../../../assets/node_modules`,免重复安装
- 同步在 `references/README.md` 注册

#### 2. 实战代码技巧文档 `references/slide-techniques.md`

从真实项目代码中沉淀的 11 个**预建模板没教**的实战模式:

- `slideConfig` 元数据对象(必加,供 compile.js 注册和 lint 校验)
- FONT 常量 + 中英文字体分离(YaHei + Arial)
- 顶部"小标签 + 章节号"全站统一模式
- **假渐变背景**(多层透明色块叠加,0 外部资源)
- 数据驱动卡片(forEach 替代硬编码)
- 大编号装饰(64pt 浅色数字)
- 对比表(colW 数组 + 末列高亮)
- 记忆点数据(`memorableData` / `memorableQuote`)
- 完整样板代码(综合应用)
- 反模式(6 条踩坑清单)
- 与 Hard Constraints 的关系

#### 3. 工具链 `scripts/`

把"做完项目后的工具"补齐,涵盖**编译前/中/后**全流程。

| 工具 | 何时跑 | 做什么 |
|------|-------|--------|
| `lint.js` | **编译前必跑** | 静态检查:NBSP、`#` 颜色、8 位 hex、async createSlide、theme 缺键 |
| `punct-fix.js` | lint 报 NBSP 时 | 全角标点 → ASCII,处理 CJK 标点 |
| `audit_pptx.py` | 完工后 | 审计元数据/主题/字体/超链接/敏感度标签,输出 JSON |
| `extract_pptx.py` | 完工后 | 抽文本(含/不含 speaker notes、按页码/范围) |
| `extract_images.py` | 完工后 | 抽图片(支持 group shape 递归) |

- 6 条 lint 规则,扫描 35 个 JS 文件通过验证
- `lint.js` 的 `SCAN_DIRS` 扩展到 `['assets', 'work', 'references/cases/real-project-example']`

### 🔧 重构 (Changed)

#### 4. `SKILL.md` 重大增强 (532 → 644 行)

**Step 4 前置:质量门禁(Hard Constraints)** 新章节,包含:

- **Hard Constraints 15 条**:从 A 真实项目沉淀的工程化红线
  - Canvas 边界、安全区域、Theme 5 键、颜色无 #、字号对比 ≥2 倍、文件命名规范等
  - 每条都标注**验证方式**和**违反后果的紧急程度**(🔴/🟡)
- **文字/图形 质量要求**:6 项自检清单
- **CJK 标点规范**:NBSP 警惕 + `punct-fix.js` 一键修复
- **工具链(scripts/)一览**:5 个工具的运行时机

**Step 4.3 Slide 模块格式**:补 `slideConfig` 必加规范

**SKILL.md 头部**:
- 20% 必查表加 `slide-techniques.md`
- 目录结构说明加 `scripts/` 5 个工具的说明
- QA 部分加 9 步完整工作流(lint → compile → audit → markitdown → extract)

**Step 4.5 Compile 脚本示例**:加 metadata 字段、OK 提示、描述性文件名

**新增"实战案例"小节**:在 22 个预建模板和 SVG 章节之间,引导用户看真实项目代码

#### 5. `assets/compile.js` 改进 (87 → 93 行)

- 新增 `createPres()` 工厂函数,集中元数据(author/title/subject)
- 添加显式 slides 数组的注释示例(替代自动发现的备选方案)
- 引入 A 的元数据最佳实践

#### 6. `references/README.md` 增强

- **新分类"⭐ 实战案例"**:明确"从理论到代码的桥梁"定位
- "按问题查找"加 2 行(`slide-techniques.md` 和 `cases/real-project-example/`)
- 更新日志加 2 条本次变更

### 🐛 修复 (Fixed)

#### 7. `assets/compile.js` 拼写

- 原始 `// compile.js - 编译脚本(增强版)` 但功能上其实是**基础版**;现在加注释区分

### 🗑️ 移除 (Removed)

无

### 🔒 安全 (Security)

#### 8. node_modules 排除

之前 `assets/node_modules`(~150MB)被提交,导致仓库膨胀。**新增 `.gitignore`**:

- 排除 `**/node_modules/`
- 排除 `output/`, `output_*/`, `*.pptx` 等编译产物
- 排除 `svg_output/`, `svg_v2/` 等中间产物
- 排除 OS / IDE / 日志 / 环境变量文件
- 排除 Trae IDE 配置

**首次使用方式**:`cd assets && npm install`(SKILL.md 已说明)

---

## 设计原则

本次变更遵循以下原则:

1. **从实战来,到实战去**:Hard Constraints 和实战技巧都是 A 真实项目踩坑后沉淀的
2. **零破坏**:所有原有 22 个预建 slide 模板、SVG 模板、案例文档**完全不动**
3. **完整闭环**:理论案例(华为/苹果) + 反面案例(failed-bp) + **实战案例(real-project-example)** 三层齐全
4. **工具链可执行**:5 个工具都经过实际运行验证(35 个 JS 文件通过 lint,真实项目 PPTX 生成成功)
5. **可逆性**:删掉新增内容即可还原(SKILL.md 还原 532 行,compile.js 还原 87 行)

---

## 如何升级现有项目

如果你有旧版本的 ppt-master,本版本不破坏任何 API,可以**直接覆盖升级**。

新用户按 SKILL.md 头部"🚀 首次使用准备"操作即可。
