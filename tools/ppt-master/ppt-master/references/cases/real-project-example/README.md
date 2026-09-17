# 真实项目案例:产品发布 PPT(10 页)

> 这是 A 的一个**真实跑通过的项目** —— 给一个叫 "MiniMax Code" 的产品做发布演示。
> 我们把项目代码原样搬过来,把所有项目特定的硬编码(品牌名/作者/输出名)改成占位符,
> 这样你既能**看到完整实战**,也能**直接 copy 改成自己的项目**。

---

## 1. 这是什么 / 不是什么

| ✅ 这是 | ❌ 这不是 |
|---------|----------|
| 一个跑得通的真实项目 | 一个抽象的"最佳实践"demo |
| 10 页完整 PPT 的代码 | 单一页的代码片段 |
| 经过项目验证的代码模式 | 教科书式的"假设性"示例 |
| 可以直接 `node compile.js` 跑 | 必须改一堆才能用 |

> 💡 **本质**:这是 [../slide-techniques.md](../slide-techniques.md) 里说的"实战代码模式"的**完整载体**。
> 看完技巧再读这个,你会有"哦原来技巧是用在这里"的实感。

---

## 2. 文件清单

```
real-project-example/
├── README.md                     ← 你正在读
├── compile.js                    ← 编译脚本（已改通用,见 §5）
├── slide-01-cover.js             ← 封面：左侧大色块 + 装饰椭圆
├── slide-02-toc.js               ← 目录：5 个章节
├── slide-03-pain-points.js       ← 痛点：2x2 网格卡片
├── slide-04-what-is.js           ← 是什么：核心定义
├── slide-05-capability-gen.js    ← 能力一：3 张堆叠卡片
├── slide-06-capability-context.js← 能力二：上下文示例
├── slide-07-capability-debug.js  ← 能力三：调试流程
├── slide-08-usecases.js          ← 5 个时间轴节点
├── slide-09-comparison.js        ← 对比表：vs 现有 AI 工具
└── slide-10-summary.js           ← 总结：3 大要点 + 关键数据
```

---

## 3. 怎么跑

```bash
# 1. 进入案例目录
cd references/cases/real-project-example

# 2. 确认 node_modules 在 B 根目录的 assets/ 下(共享依赖)
#    如果 B 根目录的 assets/node_modules 不存在,跑一次:
#    cd ../../../assets && npm install && cd -

# 3. 直接编译
node compile.js
```

预期输出:
```
OK  slide-01-cover.js
OK  slide-02-toc.js
...
OK  slide-10-summary.js

DONE: .../output/sample-product-launch-20260601.pptx
Total: 10 slides
```

---

## 4. 在每一页里能学到什么

| 页 | 文件 | 学到的实战技巧(对应 slide-techniques.md) |
|----|------|------------------------------------------|
| 1 | [slide-01-cover.js](slide-01-cover.js) | **§4 假渐变背景** —— 3 层色块叠加 + 椭圆破界 |
| 2 | [slide-02-toc.js](slide-02-toc.js) | 章节页布局 + 大编号(`05 差异化` 等) |
| 3 | [slide-03-pain-points.js](slide-03-pain-points.js) | **§5 数据驱动卡片** —— 2x2 网格 forEach |
| 4 | [slide-04-what-is.js](slide-04-what-is.js) | 大编号 `01` 当装饰(§6) |
| 5 | [slide-05-capability-gen.js](slide-05-capability-gen.js) | **§5 数据驱动卡片** + 记忆点(§8) |
| 6-7 | slide-06/07 | 同 5,展示 3 个并列能力点 |
| 8 | [slide-08-usecases.js](slide-08-usecases.js) | 时间轴(5 节点横向) |
| 9 | [slide-09-comparison.js](slide-09-comparison.js) | **§7 对比表** + 末列 `theme.accent` 高亮 |
| 10 | [slide-10-summary.js](slide-10-summary.js) | 总结页 + 大数据记忆点 |

> 💡 **每页都有这些共同结构**(见 §3 顶部小标签 + 章节号):
> - 0.3" 宽的强调色装饰条
> - "03 核心能力" 章节号标签
> - 28-30pt 的主标题
> - 13-14pt 的副标题

---

## 5. 想换成自己的项目?改 4 个地方

打开 [compile.js](compile.js),顶部注释里有完整说明。简单说改:

| # | 改什么 | 当前值 | 改成 |
|---|--------|--------|------|
| 1 | **theme 配色** | 紫蓝+粉(`6366F1` / `8B5CF6` / `F472B6`) | 你的项目配色,见 [../../color-system.md](../../color-system.md) |
| 2 | **pres.title** | `"Product Launch Demo"` | 你的演示标题 |
| 3 | **pres.author** | `"Your Name"` | 你的名字 |
| 4 | **outFile** | `sample-product-launch-${dateStr}.pptx` | `{你的主题}-${日期}.pptx`(对应 Hard Constraints #13) |

slides 数组**按需调整**:
- 删除某些页:直接从数组里注释掉
- 加新页:在数组里加 `slide-XX-xxx.js`,并创建对应文件

> ⚠️ 改 slide 内容时,**保留**这些共有结构(顶部小标签、slideConfig、FONT 常量),
> 这样整个 PPT 看起来是一套,而不是拼凑的。

---

## 6. 配套阅读

按这个顺序读,理解最顺:

1. [../../SKILL.md > Step 4 前置: 质量门禁](../../SKILL.md) —— 先知道不能踩的坑
2. [../slide-techniques.md](../slide-techniques.md) —— 11 个实战技巧清单
3. **本目录** —— 看技巧在真实项目里怎么用
4. [../../SKILL.md > Step 4: 代码生成](../../SKILL.md) —— 模板库(22 个 slide-*.js)

---

## 7. 这个项目的实际交付

A 当时编译出的 PPTX 在 A 的 `output/minimax-code-ppt-master.pptx`。
我们**不**把那个 .pptx 搬过来(项目交付物不是参考)。

如果你想看实际效果,自己跑一次 `node compile.js` 就行,10 秒内出结果。

---

## 8. 跟其他案例的关系

B 的 [references/cases/](../) 有 3 类案例:

| 类别 | 内容 | 适合 |
|------|------|------|
| 理论案例(华为/苹果/TED 拆解) | 外部产品的设计分析 | 学"什么是好设计" |
| 反面案例(failed-bp / 10-failure-patterns) | 真实失败案例 | 学"什么是不能做的" |
| **本目录(实战案例)** | **A 自己的真实项目代码** | **学"好设计怎么落地成代码"** |

**3 个层级**:理论 → 反面 → 实战。**完整闭环**。
