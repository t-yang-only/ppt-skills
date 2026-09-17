// slide-01-cover.js - 封面页
// MiniMax Code 产品发布 PPT
const pptxgen = require("pptxgenjs");

const FONT = "Microsoft YaHei";

const slideConfig = {
 type: 'cover',
 index: 1,
 title: 'MiniMax Code 产品介绍'
};

function createSlide(pres, theme) {
 const slide = pres.addSlide();
 slide.background = { color: theme.bg };

 // 左侧大色块 (模拟渐变效果: 叠三层透明色块)
 slide.addShape(pres.shapes.RECTANGLE, {
 x: 0, y: 0, w: 4.0, h: 5.625,
 fill: { color: theme.primary },
 line: { type: 'none' }
 });
 slide.addShape(pres.shapes.RECTANGLE, {
 x: 0, y: 0, w: 2.5, h: 5.625,
 fill: { color: theme.secondary, transparency: 30 },
 line: { type: 'none' }
 });
 slide.addShape(pres.shapes.OVAL, {
 x: -1.2, y: -1.0, w: 3.5, h: 3.5,
 fill: { color: theme.accent, transparency: 60 },
 line: { type: 'none' }
 });
 slide.addShape(pres.shapes.OVAL, {
 x: 1.5, y: 3.5, w: 2.8, h: 2.8,
 fill: { color: theme.light, transparency: 50 },
 line: { type: 'none' }
 });

 // 左侧色块上的小标签
 slide.addShape(pres.shapes.RECTANGLE, {
 x: 0.6, y: 0.6, w: 1.6, h: 0.04,
 fill: { color: theme.accent },
 line: { type: 'none' }
 });
 slide.addText("PRODUCT LAUNCH", {
 x: 0.6, y: 0.7, w: 3.0, h: 0.4,
 fontSize: 11,
 fontFace: "Arial",
 color: theme.light,
 bold: true,
 charSpacing: 4
 });

 // 左侧大字 (产品名英文)
 slide.addText("MiniMax", {
 x: 0.6, y: 1.5, w: 3.5, h: 0.9,
 fontSize: 56,
 fontFace: "Arial",
 color: "FFFFFF",
 bold: true
 });
 slide.addText("Code", {
 x: 0.6, y: 2.3, w: 3.5, h: 0.9,
 fontSize: 56,
 fontFace: "Arial",
 color: theme.accent,
 bold: true
 });

 // 左侧底部装饰线
 slide.addShape(pres.shapes.RECTANGLE, {
 x: 0.6, y: 3.4, w: 0.6, h: 0.04,
 fill: { color: theme.accent },
 line: { type: 'none' }
 });
 slide.addText("v1.0 GA", {
 x: 0.6, y: 3.5, w: 2.0, h: 0.3,
 fontSize: 12,
 fontFace: "Arial",
 color: theme.light
 });

 // 右侧主标题
 slide.addText("你的 AI", {
 x: 4.5, y: 1.4, w: 5.2, h: 0.9,
 fontSize: 48,
 fontFace: FONT,
 color: theme.primary,
 bold: true
 });
 slide.addText("编程伙伴", {
 x: 4.5, y: 2.2, w: 5.2, h: 0.9,
 fontSize: 48,
 fontFace: FONT,
 color: theme.primary,
 bold: true
 });

 // 副标题
 slide.addText("不止补全代码, 更是理解项目, 协同思考的智能助手", {
 x: 4.5, y: 3.3, w: 5.2, h: 0.5,
 fontSize: 16,
 fontFace: FONT,
 color: theme.secondary,
 valign: "top"
 });

 // 三个标签: 能力标签
 const tags = ["智能生成", "理解上下文", "调试测试"];
 tags.forEach((tag, i) => {
 const tx = 4.5 + i * 1.4;
 slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
 x: tx, y: 4.1, w: 1.3, h: 0.4,
 fill: { color: theme.light },
 line: { type: 'none' },
 rectRadius: 0.2
 });
 slide.addText(tag, {
 x: tx, y: 4.1, w: 1.3, h: 0.4,
 fontSize: 12,
 fontFace: FONT,
 color: theme.primary,
 align: "center",
 valign: "middle",
 bold: true
 });
 });

 // 底部信息
 slide.addShape(pres.shapes.RECTANGLE, {
 x: 4.5, y: 4.85, w: 0.4, h: 0.03,
 fill: { color: theme.accent },
 line: { type: 'none' }
 });
 slide.addText("面向开发者 / 产品发布介绍 / 2026", {
 x: 4.5, y: 4.95, w: 5.2, h: 0.3,
 fontSize: 11,
 fontFace: FONT,
 color: theme.secondary
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
 pres.writeFile({ fileName: "slide-01-preview.pptx" });
}

module.exports = { createSlide, slideConfig };
