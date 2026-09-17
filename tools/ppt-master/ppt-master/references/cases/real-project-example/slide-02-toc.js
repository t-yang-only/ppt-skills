// slide-02-toc.js - 目录页 (双栏网格)
const pptxgen = require("pptxgenjs");

const FONT = "Microsoft YaHei";

const slideConfig = {
 type: 'toc',
 index: 2,
 title: '目录'
};

function createSlide(pres, theme) {
 const slide = pres.addSlide();
 slide.background = { color: theme.bg };

 // 顶部装饰条
 slide.addShape(pres.shapes.RECTANGLE, {
 x: 0, y: 0, w: 10, h: 0.15,
 fill: { color: theme.primary },
 line: { type: 'none' }
 });
 slide.addShape(pres.shapes.RECTANGLE, {
 x: 0, y: 0, w: 3.0, h: 0.15,
 fill: { color: theme.accent },
 line: { type: 'none' }
 });

 // 页面标题
 slide.addText("CONTENTS", {
 x: 0.5, y: 0.45, w: 4.0, h: 0.4,
 fontSize: 12,
 fontFace: "Arial",
 color: theme.accent,
 bold: true,
 charSpacing: 6
 });
 slide.addText("本次分享", {
 x: 0.5, y: 0.75, w: 4.0, h: 0.7,
 fontSize: 36,
 fontFace: FONT,
 color: theme.primary,
 bold: true
 });
 slide.addText("我们想和你聊聊关于 AI 编程的那些事", {
 x: 0.5, y: 1.5, w: 4.0, h: 0.4,
 fontSize: 12,
 fontFace: FONT,
 color: theme.secondary
 });

 // 右侧装饰大数字
 slide.addText("05", {
 x: 5.5, y: 0.4, w: 4.0, h: 1.6,
 fontSize: 120,
 fontFace: "Arial",
 color: theme.light,
 bold: true,
 align: "right"
 });
 slide.addText("个核心话题", {
 x: 5.5, y: 1.6, w: 4.0, h: 0.4,
 fontSize: 13,
 fontFace: FONT,
 color: theme.secondary,
 align: "right"
 });

 // 5 个章节卡片 (左 3 + 右 2 布局)
 const chapters = [
 { num: "01", title: "开发者痛点", desc: "我们到底在为什么浪费时间?" },
 { num: "02", title: "MiniMax Code 是什么", desc: "一句话讲清楚它是谁" },
 { num: "03", title: "三大核心能力", desc: "生成 / 理解 / 调试" },
 { num: "04", title: "典型使用场景", desc: "从写代码到修 bug 全流程" },
 { num: "05", title: "和现有工具有啥不一样", desc: "为什么选 MiniMax Code" }
 ];

 // 左 3 个: 3 行 x 1 列
 chapters.slice(0, 3).forEach((ch, i) => {
 const y = 2.2 + i * 0.95;
 // 编号圆点
 slide.addShape(pres.shapes.OVAL, {
 x: 0.5, y: y + 0.1, w: 0.55, h: 0.55,
 fill: { color: theme.primary },
 line: { type: 'none' }
 });
 slide.addText(ch.num, {
 x: 0.5, y: y + 0.1, w: 0.55, h: 0.55,
 fontSize: 14,
 fontFace: "Arial",
 color: "FFFFFF",
 bold: true,
 align: "center",
 valign: "middle"
 });
 // 标题
 slide.addText(ch.title, {
 x: 1.25, y: y + 0.05, w: 3.5, h: 0.35,
 fontSize: 16,
 fontFace: FONT,
 color: theme.primary,
 bold: true
 });
 // 描述
 slide.addText(ch.desc, {
 x: 1.25, y: y + 0.42, w: 3.5, h: 0.3,
 fontSize: 11,
 fontFace: FONT,
 color: theme.secondary
 });
 // 分割线
 if (i < 2) {
 slide.addShape(pres.shapes.RECTANGLE, {
 x: 1.25, y: y + 0.85, w: 3.3, h: 0.01,
 fill: { color: theme.light },
 line: { type: 'none' }
 });
 }
 });

 // 右 2 个: 2 行
 chapters.slice(3).forEach((ch, i) => {
 const y = 2.2 + i * 0.95;
 // 编号圆点
 slide.addShape(pres.shapes.OVAL, {
 x: 5.5, y: y + 0.1, w: 0.55, h: 0.55,
 fill: { color: theme.accent },
 line: { type: 'none' }
 });
 slide.addText(ch.num, {
 x: 5.5, y: y + 0.1, w: 0.55, h: 0.55,
 fontSize: 14,
 fontFace: "Arial",
 color: "FFFFFF",
 bold: true,
 align: "center",
 valign: "middle"
 });
 // 标题
 slide.addText(ch.title, {
 x: 6.25, y: y + 0.05, w: 3.5, h: 0.35,
 fontSize: 16,
 fontFace: FONT,
 color: theme.primary,
 bold: true
 });
 // 描述
 slide.addText(ch.desc, {
 x: 6.25, y: y + 0.42, w: 3.5, h: 0.3,
 fontSize: 11,
 fontFace: FONT,
 color: theme.secondary
 });
 if (i < 1) {
 slide.addShape(pres.shapes.RECTANGLE, {
 x: 6.25, y: y + 0.85, w: 3.3, h: 0.01,
 fill: { color: theme.light },
 line: { type: 'none' }
 });
 }
 });

 // 底部装饰
 slide.addShape(pres.shapes.RECTANGLE, {
 x: 0, y: 5.475, w: 10, h: 0.15,
 fill: { color: theme.light },
 line: { type: 'none' }
 });
 slide.addText("02", {
 x: 9.0, y: 5.1, w: 0.5, h: 0.3,
 fontSize: 10,
 fontFace: "Arial",
 color: theme.secondary,
 align: "right"
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
 pres.writeFile({ fileName: "slide-02-preview.pptx" });
}

module.exports = { createSlide, slideConfig };
