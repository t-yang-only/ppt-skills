// compile.js - 真实项目示例的编译脚本
// =====================================
// 用途:把当前目录的 10 个 slide-XX.js 编译成 PPTX
// 来源:这是 A 真实项目(MiniMax Code 产品发布)的代码
// 学习点:见本目录 README.md,以及 ../slide-techniques.md
//
// 跑法:
//   cd references/cases/real-project-example
//   node compile.js
//
// 想换成自己的项目?改 4 个地方:
//   1. theme 配色    → 改成你项目的颜色
//   2. pres.title    → 你的演示标题
//   3. pres.author   → 你的名字
//   4. outFile       → 你的输出文件名
// =====================================
const path = require("path");
const fs = require("fs");

// 智能加载 pptxgenjs:
// 1) 优先从父目录的 node_modules 找（ppt-master 标准结构）
// 2) 找不到时回退到 B 的 assets/node_modules（开发期）
// 3) 再不行就回退到相对路径（用户复制本目录到独立项目时）
let pptxgen;
try {
  pptxgen = require("pptxgenjs");
} catch (e1) {
  try {
    pptxgen = require("../../../assets/node_modules/pptxgenjs");
  } catch (e2) {
    pptxgen = require("pptxgenjs");  // 最后一次尝试,让错误信息更清晰
  }
}

// ---- ① Theme 配色（5 键结构,见 SKILL.md 4.2 节） ----
// 这是"现代产品发布"风格的紫蓝+粉配色,作为示例
// 实际项目请按 references/color-system.md 选择配色
const theme = {
  primary: "6366F1",    // Indigo 紫蓝 - 主色（标题/深背景）
  secondary: "8B5CF6",  // Violet 紫 - 次要（副标题）
  accent: "F472B6",     // Pink 粉 - 强调（记忆点/按钮）
  light: "EEF2FF",      // 浅紫 - 装饰（卡片底）
  bg: "FFFFFF"          // 白 - 背景
};

// ---- ② 输出目录(本目录下的 output/,不要污染父目录) ----
const outputDir = path.join(__dirname, "output");
if (!fs.existsSync(outputDir)) {
  fs.mkdirSync(outputDir, { recursive: true });
}

// ---- ③ 创建 PPT 实例 + metadata(对应 Hard Constraints #1, #13) ----
const pres = new pptxgen();
pres.layout = "LAYOUT_16x9";
pres.author = "Your Name";                // ← 改成你的名字
pres.title = "Product Launch Demo";       // ← 改成你的标题
pres.subject = "Generated with PptxGenJS";

// ---- ④ 显式 slides 数组(替代自动发现,顺序可控) ----
const slides = [
  "slide-01-cover.js",
  "slide-02-toc.js",
  "slide-03-pain-points.js",
  "slide-04-what-is.js",
  "slide-05-capability-gen.js",
  "slide-06-capability-context.js",
  "slide-07-capability-debug.js",
  "slide-08-usecases.js",
  "slide-09-comparison.js",
  "slide-10-summary.js"
];

// ---- ⑤ 加载并生成(单页失败不中断) ----
let slideCount = 0;
slides.forEach((file) => {
  try {
    const mod = require(path.join(__dirname, file));
    mod.createSlide(pres, theme);
    slideCount++;
    console.log(`OK  ${file}`);
  } catch (err) {
    console.error(`FAIL  ${file}: ${err.message}`);
  }
});

// ---- ⑥ 输出文件名(描述性:主题-日期.pptx) ----
const dateStr = new Date().toISOString().slice(0, 10).replace(/-/g, "");
const outFile = path.join(outputDir, `sample-product-launch-${dateStr}.pptx`);

pres.writeFile({ fileName: outFile })
  .then(() => {
    console.log(`\nDONE: ${outFile}`);
    console.log(`Total: ${slideCount} slides`);
  })
  .catch((err) => {
    console.error("Write failed:", err);
  });
