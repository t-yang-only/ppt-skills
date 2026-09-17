// compile.js - 编译脚本（增强版）
// 功能：--watch 监听模式 | --preview 单页预览 | --list 列出页面
//
// ✅ 学习自 A 的 work/minimax-code/compile.js 的最佳实践：
//   1. 显式 metadata 字段（author / title / subject）
//   2. path.join 跨平台路径
//   3. 单页失败不中断（try/catch per slide）
//
// 💡 显式 slides 数组（替代自动发现，适合"挑页"或"排页"场景）：
//   const slides = [
//     "slide-01-cover.js",
//     "slide-02-toc.js",
//     // 注释掉不想生成的页: "slide-03-section.js",
//   ];
const pptxgen = require("pptxgenjs");
const fs = require("fs");
const path = require("path");

const args = process.argv.slice(2);
const mode = args.includes("--watch") ? "watch" : args.includes("--preview") ? "preview" : "build";

const theme = {
  primary: "1A237E",
  secondary: "3949AB",
  accent: "F57C00",
  light: "E8EAF6",
  bg: "FFFFFF"
};

function createPres() {
  const pres = new pptxgen();
  pres.layout = "LAYOUT_16x9";
  pres.author = "PPT Master";
  pres.title = "演示文稿";
  pres.subject = "Generated with PptxGenJS";
  return pres;
}

const outputDir = path.join(__dirname, "output");
if (!fs.existsSync(outputDir)) fs.mkdirSync(outputDir, { recursive: true });

const slides = fs.readdirSync(__dirname)
  .filter(f => f.startsWith("slide-") && f.endsWith(".js"))
  .sort();

if (mode === "list") {
  console.log("可用页面：");
  slides.forEach((f, i) => console.log(`  ${i + 1}. ${f}`));
  process.exit(0);
}

function buildOne(slideFile) {
  const pres = createPres();
  try {
    const mod = require(path.join(__dirname, slideFile));
    mod.createSlide(pres, theme, {});
    const base = slideFile.replace(".js", "");
    pres.writeFile({ fileName: path.join(outputDir, `${base}.pptx`) })
      .then(() => console.log(`✓ ${base}.pptx`))
      .catch(err => console.error(`✗ ${base}: ${err.message}`));
  } catch (err) {
    console.error(`✗ ${slideFile}: ${err.message}`);
  }
}

function buildAll() {
  const pres = createPres();
  let ok = 0, fail = 0;
  slides.forEach(f => {
    try {
      const mod = require(path.join(__dirname, f));
      mod.createSlide(pres, theme, {});
      console.log(`✓ ${f}`);
      ok++;
    } catch (err) {
      console.error(`✗ ${f}: ${err.message}`);
      fail++;
    }
  });
  console.log(`\n编译完成: ${ok} 成功, ${fail} 失败`);
  const out = path.join(outputDir, "presentation.pptx");
  pres.writeFile({ fileName: out })
    .then(() => console.log(`✓ 生成: ${out}`))
    .catch(err => console.error(`✗ 失败: ${err.message}`));
}

if (mode === "preview") {
  const idx = parseInt(args[args.indexOf("--preview") + 1]) - 1;
  if (isNaN(idx) || idx < 0 || idx >= slides.length) {
    console.error(`无效页码，可用: 1-${slides.length}`);
    process.exit(1);
  }
  console.log(`预览第 ${idx + 1} 页: ${slides[idx]}`);
  buildOne(slides[idx]);
} else if (mode === "watch") {
  console.log(`监听中 (${slides.length} 个页面)，修改后自动重建...\n`);
  buildAll();
  fs.watch(__dirname, { persistent: true }, (evt, f) => {
    if (f && f.startsWith("slide-") && f.endsWith(".js")) {
      console.log(`\n▸ 检测到变化: ${f}`);
      buildAll();
    }
  });
} else {
  buildAll();
}
