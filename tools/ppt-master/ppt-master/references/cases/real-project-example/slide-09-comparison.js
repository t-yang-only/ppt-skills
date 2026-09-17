// slide-09-comparison.js - 与现有 AI 工具的差异
const pptxgen = require("pptxgenjs");

const FONT = "Microsoft YaHei";

const slideConfig = {
 type: 'content',
 index: 9,
 title: '与现有 AI 工具的差异'
};

function createSlide(pres, theme) {
 const slide = pres.addSlide();
 slide.background = { color: theme.bg };

 // 顶部小标签
 slide.addShape(pres.shapes.RECTANGLE, {
 x: 0.5, y: 0.4, w: 0.3, h: 0.04,
 fill: { color: theme.accent },
 line: { type: 'none' }
 });
 slide.addText("05 差异化", {
 x: 0.85, y: 0.3, w: 4.0, h: 0.3,
 fontSize: 11,
 fontFace: "Arial",
 color: theme.accent,
 bold: true,
 charSpacing: 3
 });

 // 主标题
 slide.addText("和现有 AI 工具到底有啥不一样", {
 x: 0.5, y: 0.7, w: 9.0, h: 0.7,
 fontSize: 28,
 fontFace: FONT,
 color: theme.primary,
 bold: true
 });

 // 对比表头
 const tableX = 0.5;
 const tableY = 1.6;
 const colW = [2.0, 2.3, 2.3, 2.4];
 const tableW = colW.reduce((a, b) => a + b, 0);

 // 表头背景
 slide.addShape(pres.shapes.RECTANGLE, {
 x: tableX, y: tableY, w: tableW, h: 0.5,
 fill: { color: theme.primary },
 line: { type: 'none' }
 });

 const headers = ["维度", "代码补全类\n(Copilot 等)", "Chat 类\n(ChatGPT/Claude 等)", "MiniMax Code"];
 let hx = tableX;
 headers.forEach((h, i) => {
 slide.addText(h, {
 x: hx, y: tableY, w: colW[i], h: 0.5,
 fontSize: 11,
 fontFace: FONT,
 color: i === 3 ? theme.accent : "FFFFFF",
 bold: true,
 align: "center",
 valign: "middle"
 });
 hx += colW[i];
 });

 // 表头右侧色块强调
 slide.addShape(pres.shapes.RECTANGLE, {
 x: tableX + colW[0] + colW[1] + colW[2], y: tableY, w: 0.06, h: 0.5,
 fill: { color: theme.accent },
 line: { type: 'none' }
 });

 // 5 个对比行
 const rows = [
 {
 dim: "上下文范围",
 a: "当前文件 / 几行",
 b: "整段 prompt",
 c: "整个项目 + 历史"
 },
 {
 dim: "工作方式",
 a: "逐行补全",
 b: "对话式问答",
 c: "Agentic 多步执行"
 },
 {
 dim: "工具调用",
 a: "无",
 b: "插件 / 需手动",
 c: "原生内置 (终端, Git, 测试)"
 },
 {
 dim: "代码质量",
 a: "补全即用, 无审查",
 b: "建议式, 需人工落地",
 c: "可审查 diff + 测试 + 文档"
 },
 {
 dim: "私有部署",
 a: "云端为主",
 b: "云端为主",
 c: "本地 + 私有化均可"
 }
 ];

 const rowH = 0.55;
 rows.forEach((r, i) => {
 const y = tableY + 0.5 + i * rowH;
 // 斑马纹
 if (i % 2 === 0) {
 slide.addShape(pres.shapes.RECTANGLE, {
 x: tableX, y: y, w: tableW, h: rowH,
 fill: { color: theme.light, transparency: 50 },
 line: { type: 'none' }
 });
 }
 // MiniMax Code 列高亮
 slide.addShape(pres.shapes.RECTANGLE, {
 x: tableX + colW[0] + colW[1] + colW[2], y: y, w: colW[3], h: rowH,
 fill: { color: theme.accent, transparency: 85 },
 line: { type: 'none' }
 });

 // 维度
 slide.addText(r.dim, {
 x: tableX, y: y, w: colW[0], h: rowH,
 fontSize: 11,
 fontFace: FONT,
 color: theme.primary,
 bold: true,
 align: "center",
 valign: "middle"
 });
 // A 列
 slide.addText(r.a, {
 x: tableX + colW[0], y: y, w: colW[1], h: rowH,
 fontSize: 10,
 fontFace: FONT,
 color: theme.secondary,
 align: "center",
 valign: "middle"
 });
 // B 列
 slide.addText(r.b, {
 x: tableX + colW[0] + colW[1], y: y, w: colW[2], h: rowH,
 fontSize: 10,
 fontFace: FONT,
 color: theme.secondary,
 align: "center",
 valign: "middle"
 });
 // C 列 (我们)
 slide.addText(r.c, {
 x: tableX + colW[0] + colW[1] + colW[2], y: y, w: colW[3], h: rowH,
 fontSize: 10.5,
 fontFace: FONT,
 color: theme.primary,
 bold: true,
 align: "center",
 valign: "middle"
 });
 });

 // 底部记忆点
 slide.addShape(pres.shapes.RECTANGLE, {
 x: 0.5, y: 5.2, w: 9.0, h: 0.04,
 fill: { color: theme.accent },
 line: { type: 'none' }
 });
 slide.addText("一句话差异 / Copilot 帮你打字, ChatGPT 陪你聊天, MiniMax Code 帮你把活干完", {
 x: 0.5, y: 5.3, w: 9.0, h: 0.3,
 fontSize: 11,
 fontFace: FONT,
 color: theme.accent,
 bold: true
 });

 return slide;
}

if (require.main === module) {
 const pres = new pptxgen();
 pres.layout = 'LAYOUT_16x9';
 const theme = {
 primary: "6366F1", secondary: "8B5CF6", accent: "F472B6",
 light: "EEF2FF", bg: "FFFFFF"
 };
 createSlide(pres, theme);
 pres.writeFile({ fileName: "slide-09-preview.pptx" });
}

module.exports = { createSlide, slideConfig };
