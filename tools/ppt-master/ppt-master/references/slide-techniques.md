# 实战幻灯片技巧(Real-world Slide Techniques)

> 来源:从 `ppt-master` 的真实项目代码中提炼(参见 `assets/slide-01-cover.js` 及相关 work/ 项目)。
> 这些是**预建模板没教但实战中必备**的代码模式,直接抄过去用。

---

## 1. slideConfig 元数据对象(必加)

每个 slide 文件**必须**在顶部声明 `slideConfig`,供 `compile.js` 做注册中心、TOC 自动生成、lint 校验。

```javascript
const slideConfig = {
  type: 'cover',       // cover | toc | section | content | summary | end
  index: 1,            // 页码
  title: 'MiniMax Code 产品介绍'
};

function createSlide(pres, theme, options = {}) {
  // ...
}

module.exports = { createSlide, slideConfig };
```

**收益**:
- 自动生成 TOC 不用手填
- compile.js 可以按 type 分组/排序
- lint 可以按 type 检查必备字段

---

## 2. FONT 常量 + 中英文字体分离(强制)

```javascript
const FONT = "Microsoft YaHei";   // 中文
const FONT_EN = "Arial";          // 英文/数字

// ✅ 正确:中文用 YaHei,英文用 Arial
slide.addText("产品发布", { fontFace: FONT, fontSize: 30 });
slide.addText("PRODUCT LAUNCH", { fontFace: FONT_EN, fontSize: 11, charSpacing: 4 });

// ✗ 错误:全文用同一字体,英文会显得呆板
slide.addText("PRODUCT LAUNCH 产品发布", { fontFace: FONT });
```

**为什么分开**:中英文混排时,YaHei 的英文部分会显得"胖而圆",Arial 笔画更利落。混排时分别指定字体,观感立刻专业。

---

## 3. 顶部"小标签 + 章节号"模式(全站统一)

A 真实项目里**每页都重复**的固定结构:

```javascript
// 1) 左侧 0.3" 宽的强调色装饰条
slide.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 0.4, w: 0.3, h: 0.04,
  fill: { color: theme.accent },
  line: { type: 'none' }
});

// 2) 装饰条右侧的章节标签
slide.addText("01 痛点场景", {
  x: 0.85, y: 0.3, w: 4.0, h: 0.3,
  fontSize: 11,
  fontFace: FONT_EN,
  color: theme.accent,
  bold: true,
  charSpacing: 3
});
```

**为什么有效**:
- 观众一眼看出"当前在第几章"
- 0.3" 的小红条是低成本高回报的视觉锚点
- `charSpacing: 3` 让英文标签显得更"高级"

---

## 4. 假渐变背景(零成本出彩)

用**多层透明色块叠加**模拟渐变效果,不需要图片:

```javascript
// 底层:大色块
slide.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 4.0, h: 5.625,
  fill: { color: theme.primary },
  line: { type: 'none' }
});

// 中层:叠一个稍窄的色块,加 transparency
slide.addShape(pres.shapes.RECTANGLE, {
  x: 0, y: 0, w: 2.5, h: 5.625,
  fill: { color: theme.secondary, transparency: 30 },
  line: { type: 'none' }
});

// 装饰:溢出边界的椭圆
slide.addShape(pres.shapes.OVAL, {
  x: -1.2, y: -1.0, w: 3.5, h: 3.5,
  fill: { color: theme.accent, transparency: 60 },
  line: { type: 'none' }
});
```

**技巧**:
- `transparency: 30` 表示 30% 透明(0=不透明,100=全透明)
- 椭圆可以**故意溢出边界**(`x: -1.2`),制造"破界"的高级感
- 3 层叠加就有"渐变"质感,但 0 外部资源依赖

---

## 5. 数据驱动卡片(避免硬编码)

不要每张卡片手动复制代码,用 `forEach` + 数据数组:

```javascript
// ✅ 推荐:数据驱动
const cards = [
  { icon: "💡", title: "创新思维", desc: "突破传统框架" },
  { icon: "📊", title: "数据驱动", desc: "基于市场数据" },
  { icon: "🚀", title: "快速执行", desc: "敏捷开发" }
];

cards.forEach((card, i) => {
  const x = 0.5 + i * 3.0;  // 横向排列
  slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
    x: x, y: 2.0, w: 2.8, h: 2.5,
    fill: { color: theme.light },
    line: { type: 'none' }
  });
  slide.addText(card.icon, {
    x: x + 0.2, y: 2.2, w: 0.6, h: 0.6,
    fontSize: 32
  });
  slide.addText(card.title, {
    x: x + 0.2, y: 2.9, w: 2.4, h: 0.5,
    fontSize: 18, bold: true, color: theme.primary
  });
  slide.addText(card.desc, {
    x: x + 0.2, y: 3.4, w: 2.4, h: 0.8,
    fontSize: 12, color: theme.secondary
  });
});

// ✗ 错误:3 张卡片手写 3 遍
// (维护噩梦,改一处要改 3 处)
```

---

## 6. 大编号技巧(数字本身就是装饰)

A 的章节页用 **64pt 的浅色大数字**当装饰元素:

```javascript
slide.addText("01", {
  x: 0.5, y: 0.6, w: 1.2, h: 0.9,
  fontSize: 64,
  fontFace: FONT_EN,
  color: theme.light,   // 浅色,作为底纹
  bold: true
});
```

**为什么用浅色**:不抢主标题的视觉权重,但填充了左上角空白,让版面"满"起来。

---

## 7. 对比表(用 colW 数组)

```javascript
const colW = [2.0, 2.3, 2.3, 2.4];   // 4 列宽度
const tableW = colW.reduce((a, b) => a + b, 0);  // 9.0

// 表头背景
slide.addShape(pres.shapes.RECTANGLE, {
  x: 0.5, y: 1.6, w: tableW, h: 0.5,
  fill: { color: theme.primary },
  line: { type: 'none' }
});

// 表头文字(累加 x 坐标)
const headers = ["维度", "代码补全类", "Chat 类", "MiniMax Code"];
let hx = 0.5;
headers.forEach((h, i) => {
  slide.addText(h, {
    x: hx, y: 1.6, w: colW[i], h: 0.5,
    fontSize: 11,
    color: i === 3 ? theme.accent : "FFFFFF",  // 末列高亮
    bold: true
  });
  hx += colW[i];
});
```

**技巧**:末列用 `theme.accent` 高亮,强调"我们的方案"。

---

## 8. 记忆点数据(memorableData / memorableQuote)

A 项目里**几乎每张内容页都有 `memorableData` 或 `memorableQuote`**:

```javascript
memorableData: "客户满意度提升 47%"  // 数字型记忆点
memorableQuote: "携手共创,合作共赢"  // 金句型记忆点
memorableEvent: "关键里程碑:A 轮融资" // 事件型记忆点
```

**为什么**:`design-principles.md` 的 Q4 要求每页必须有"记忆点",在代码里直接以字段形式承载,既便于 compile.js 抽出来做"金句页",也防止遗漏。

---

## 9. 完整样板(综合应用)

把上述 8 条技巧组合起来,就是 A 真实项目的标准 slide 文件结构:

```javascript
// slide-XX-xxx.js - 简述这页内容
const pptxgen = require("pptxgenjs");

// ① 字体常量
const FONT = "Microsoft YaHei";
const FONT_EN = "Arial";

// ② slideConfig 元数据
const slideConfig = {
  type: 'content',
  index: 5,
  title: '能力一:智能代码生成'
};

function createSlide(pres, theme, options = {}) {
  const slide = pres.addSlide();
  slide.background = { color: theme.bg };

  // ③ 顶部"小标签 + 章节号"
  slide.addShape(pres.shapes.RECTANGLE, {
    x: 0.5, y: 0.4, w: 0.3, h: 0.04,
    fill: { color: theme.accent },
    line: { type: 'none' }
  });
  slide.addText("03 核心能力", {
    x: 0.85, y: 0.3, w: 4.0, h: 0.3,
    fontSize: 11, fontFace: FONT_EN,
    color: theme.accent, bold: true, charSpacing: 3
  });

  // ④ 大编号装饰
  slide.addText("01", {
    x: 0.5, y: 0.6, w: 1.2, h: 0.9,
    fontSize: 64, fontFace: FONT_EN,
    color: theme.light, bold: true
  });

  // ⑤ 主标题 + 副标题
  slide.addText(options.title || "智能代码生成与补全", {
    x: 1.6, y: 0.8, w: 7.0, h: 0.6,
    fontSize: 30, fontFace: FONT,
    color: theme.primary, bold: true
  });
  slide.addText(options.subtitle || "从「按行补全」到「按意图生成」", {
    x: 1.6, y: 1.4, w: 7.5, h: 0.4,
    fontSize: 13, fontFace: FONT, color: theme.secondary
  });

  // ⑥ 数据驱动卡片
  const points = options.points || [];
  points.forEach((p, i) => {
    const y = 2.1 + i * 1.05;
    slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
      x: 0.5, y: y, w: 4.4, h: 0.95,
      fill: { color: theme.light },
      line: { type: 'none' }
    });
    slide.addText(p.label, {
      x: 0.7, y: y + 0.1, w: 4.0, h: 0.4,
      fontSize: 16, fontFace: FONT, color: theme.primary, bold: true
    });
    slide.addText(p.desc, {
      x: 0.7, y: y + 0.5, w: 4.0, h: 0.4,
      fontSize: 11, fontFace: FONT, color: theme.secondary
    });
  });

  // ⑧ 记忆点(放右下角,大字)
  if (options.memorableData) {
    slide.addText(options.memorableData, {
      x: 5.5, y: 4.6, w: 4.0, h: 0.6,
      fontSize: 22, fontFace: FONT_EN,
      color: theme.accent, bold: true
    });
  }
}

// ⑨ 导出(包含 slideConfig)
module.exports = { createSlide, slideConfig };
```

---

## 10. 反模式(避免)

| 反模式 | 后果 | 修正 |
|--------|------|------|
| 颜色硬编码 `"1A237E"` | 改主题要全局替换 | 用 `theme.primary` |
| 字号乱选 `12/14/18/22/26` | 对比不足 | 限定在 FONT 常量里 |
| 卡片手写 N 遍 | 维护噩梦 | 数据驱动 + forEach |
| 没有 slideConfig | compile.js 不知道是封面还是内容页 | 必加 |
| 中英文混用 YaHei | 英文显得呆 | 分别指定 FONT / FONT_EN |
| 装饰元素超出安全区 | 投影时裁剪 | x∈[0.5,9.5], y∈[0.4,5.225] |

---

## 11. 与 Hard Constraints 的关系

> 上面的所有技巧都**不违反** [SKILL.md > Step 4 > Hard Constraints](../SKILL.md):
> - 安全区域 (rule #3)
> - 字号对比 ≥ 2 倍 (rule #10)
> - 颜色不超过 3 种 (rule #11)
> - 复用 option 对象用工厂 (rule #12)
>
> 实际写代码时,把这两份文档**对照着看**,可以同时满足"工程化"和"美学"两个维度的要求。
