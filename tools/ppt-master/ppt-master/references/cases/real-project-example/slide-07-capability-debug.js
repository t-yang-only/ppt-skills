// slide-07-capability-debug.js - 能力 3: 调试与测试
const pptxgen = require("pptxgenjs");

const FONT = "Microsoft YaHei";

const slideConfig = {
 type: 'content',
 index: 7,
 title: '能力三:调试与测试'
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
 slide.addText("03 核心能力", {
 x: 0.85, y: 0.3, w: 4.0, h: 0.3,
 fontSize: 11,
 fontFace: "Arial",
 color: theme.accent,
 bold: true,
 charSpacing: 3
 });

 // 大编号
 slide.addText("03", {
 x: 0.5, y: 0.6, w: 1.2, h: 0.9,
 fontSize: 64,
 fontFace: "Arial",
 color: theme.light,
 bold: true
 });

 // 主标题
 slide.addText("调试, 测试与质量保障", {
 x: 1.6, y: 0.8, w: 7.0, h: 0.6,
 fontSize: 30,
 fontFace: FONT,
 color: theme.primary,
 bold: true
 });
 slide.addText("不止是「写代码」, 更是帮你「少写 bug」「少返工」", {
 x: 1.6, y: 1.4, w: 7.5, h: 0.4,
 fontSize: 13,
 fontFace: FONT,
 color: theme.secondary
 });

 // 4 个能力卡片 (2x2 网格)
 const items = [
 {
 icon: "B",
 title: "Bug 定位助手",
 desc: "贴入错误堆栈 / 报错日志, AI 自动分析根因\n给出修复建议和最小可行 patch",
 tag: "DEBUG"
 },
 {
 icon: "T",
 title: "自动生成单测",
 desc: "选中函数即可生成边界用例 + Mock\n覆盖正常, 异常, 性能多维度",
 tag: "TEST"
 },
 {
 icon: "R",
 title: "智能 Code Review",
 desc: "在 PR 阶段自动审查, 标出潜在问题\n风格不一致, 并发风险, 性能瓶颈",
 tag: "REVIEW"
 },
 {
 icon: "Q",
 title: "质量度量",
 desc: "为每个模块生成质量报告\n复杂圈数, 重复率, 可维护性指数",
 tag: "METRIC"
 }
 ];

 const cardW = 4.35;
 const cardH = 1.55;
 const startX = 0.5;
 const startY = 1.95;
 const gapX = 0.3;
 const gapY = 0.2;

 items.forEach((it, i) => {
 const col = i % 2;
 const row = Math.floor(i / 2);
 const x = startX + col * (cardW + gapX);
 const y = startY + row * (cardH + gapY);

 // 卡片底
 slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
 x: x, y: y, w: cardW, h: cardH,
 fill: { color: theme.light },
 line: { type: 'none' },
 rectRadius: 0.1
 });

 // 字母圆形
 slide.addShape(pres.shapes.OVAL, {
 x: x + 0.3, y: y + 0.3, w: 0.7, h: 0.7,
 fill: { color: theme.primary },
 line: { type: 'none' }
 });
 slide.addText(it.icon, {
 x: x + 0.3, y: y + 0.3, w: 0.7, h: 0.7,
 fontSize: 28,
 fontFace: "Arial",
 color: "FFFFFF",
 bold: true,
 align: "center",
 valign: "middle"
 });

 // 标题
 slide.addText(it.title, {
 x: x + 1.15, y: y + 0.3, w: cardW - 1.8, h: 0.4,
 fontSize: 17,
 fontFace: FONT,
 color: theme.primary,
 bold: true,
 valign: "middle"
 });

 // tag
 slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
 x: x + cardW - 0.7, y: y + 0.35, w: 0.5, h: 0.28,
 fill: { color: theme.accent },
 line: { type: 'none' },
 rectRadius: 0.04
 });
 slide.addText(it.tag, {
 x: x + cardW - 0.7, y: y + 0.35, w: 0.5, h: 0.28,
 fontSize: 8,
 fontFace: "Arial",
 color: "FFFFFF",
 bold: true,
 align: "center",
 valign: "middle"
 });

 // 描述
 slide.addText(it.desc, {
 x: x + 1.15, y: y + 0.75, w: cardW - 1.3, h: 0.7,
 fontSize: 11,
 fontFace: FONT,
 color: theme.secondary,
 valign: "top"
 });
 });

 // 底部记忆点
 slide.addShape(pres.shapes.RECTANGLE, {
 x: 0.5, y: 5.3, w: 9.0, h: 0.04,
 fill: { color: theme.accent },
 line: { type: 'none' }
 });
 slide.addText("亮点 / 从「写完代码就算完事」升级为「带着质量门禁的完整工程闭环」", {
 x: 0.5, y: 5.4, w: 9.0, h: 0.3,
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
 pres.writeFile({ fileName: "slide-07-preview.pptx" });
}

module.exports = { createSlide, slideConfig };
