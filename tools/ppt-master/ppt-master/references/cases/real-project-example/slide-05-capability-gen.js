// slide-05-capability-gen.js - 能力 1: 智能代码生成
const pptxgen = require("pptxgenjs");

const FONT = "Microsoft YaHei";

const slideConfig = {
 type: 'content',
 index: 5,
 title: '能力一:智能代码生成'
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
 slide.addText("01", {
 x: 0.5, y: 0.6, w: 1.2, h: 0.9,
 fontSize: 64,
 fontFace: "Arial",
 color: theme.light,
 bold: true
 });

 // 主标题
 slide.addText("智能代码生成与补全", {
 x: 1.6, y: 0.8, w: 7.0, h: 0.6,
 fontSize: 30,
 fontFace: FONT,
 color: theme.primary,
 bold: true
 });
 slide.addText("从「按行补全」到「按意图生成」, 从「写一句」到「建一块」", {
 x: 1.6, y: 1.4, w: 7.5, h: 0.4,
 fontSize: 13,
 fontFace: FONT,
 color: theme.secondary
 });

 // 左侧 3 个能力点
 const points = [
 {
 label: "行内补全",
 desc: "秒级响应, 根据光标位置和上下文智能预测\n下一个 token, 准确率领先"
 },
 {
 label: "块级生成",
 desc: "选中一段描述, AI 自动写出函数 / 类 / 模块\n支持多文件联动"
 },
 {
 label: "从零到一",
 desc: "一句自然语言 = 一个完整功能\n从脚手架到测试用例一气呵成"
 }
 ];

 const pX = 0.5;
 const pStartY = 2.1;
 const pH = 0.95;
 const pGap = 0.1;

 points.forEach((p, i) => {
 const y = pStartY + i * (pH + pGap);
 // 卡片
 slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
 x: pX, y: y, w: 4.4, h: pH,
 fill: { color: theme.light },
 line: { type: 'none' },
 rectRadius: 0.08
 });
 // 左侧色条
 slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
 x: pX, y: y, w: 0.08, h: pH,
 fill: { color: theme.accent },
 line: { type: 'none' }
 });
 // 标题
 slide.addText(p.label, {
 x: pX + 0.3, y: y + 0.1, w: 4.0, h: 0.35,
 fontSize: 16,
 fontFace: FONT,
 color: theme.primary,
 bold: true
 });
 // 描述
 slide.addText(p.desc, {
 x: pX + 0.3, y: y + 0.45, w: 4.0, h: 0.5,
 fontSize: 10,
 fontFace: FONT,
 color: theme.secondary,
 valign: "top"
 });
 });

 // 右侧: 代码示例卡片
 const codeX = 5.2;
 const codeY = 2.1;
 const codeW = 4.3;
 const codeH = 3.05;

 // 卡片背景
 slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
 x: codeX, y: codeY, w: codeW, h: codeH,
 fill: { color: "1E1B4B" },
 line: { type: 'none' },
 rectRadius: 0.1
 });

 // 模拟 mac 窗口红绿灯
 slide.addShape(pres.shapes.OVAL, {
 x: codeX + 0.2, y: codeY + 0.18, w: 0.13, h: 0.13,
 fill: { color: "F472B6" },
 line: { type: 'none' }
 });
 slide.addShape(pres.shapes.OVAL, {
 x: codeX + 0.4, y: codeY + 0.18, w: 0.13, h: 0.13,
 fill: { color: "FBBF24" },
 line: { type: 'none' }
 });
 slide.addShape(pres.shapes.OVAL, {
 x: codeX + 0.6, y: codeY + 0.18, w: 0.13, h: 0.13,
 fill: { color: "34D399" },
 line: { type: 'none' }
 });
 slide.addText("user.py", {
 x: codeX + 1.0, y: codeY + 0.1, w: 2.0, h: 0.3,
 fontSize: 10,
 fontFace: "Arial",
 color: theme.light
 });

 // 代码行
 const codeLines = [
 { text: "def ", color: "F472B6" },
 { text: "create_user", color: "60A5FA" },
 { text: "(", color: "E0E7FF" },
 { text: "name", color: "FBBF24" },
 { text: ", ", color: "E0E7FF" },
 { text: "email", color: "FBBF24" },
 { text: "):", color: "E0E7FF" }
 ];
 // 第 1 行
 let cx = codeX + 0.25;
 codeLines.forEach((seg) => {
 slide.addText(seg.text, {
 x: cx, y: codeY + 0.5, w: 2.5, h: 0.3,
 fontSize: 13,
 fontFace: "Courier New",
 color: seg.color
 });
 cx += seg.text.length * 0.085;
 });

 // 用户注释
 slide.addText(" # 创建一个新用户", {
 x: codeX + 0.25, y: codeY + 0.85, w: 4.0, h: 0.3,
 fontSize: 12,
 fontFace: "Courier New",
 color: "6B7280"
 });

 // AI 生成的代码 (高亮)
 slide.addShape(pres.shapes.RECTANGLE, {
 x: codeX + 0.25, y: codeY + 1.2, w: 3.8, h: 1.7,
 fill: { color: "6366F1", transparency: 70 },
 line: { type: 'none' }
 });
 slide.addText(" user = User(", {
 x: codeX + 0.3, y: codeY + 1.25, w: 4.0, h: 0.3,
 fontSize: 12,
 fontFace: "Courier New",
 color: "E0E7FF"
 });
 slide.addText(" name=name,", {
 x: codeX + 0.3, y: codeY + 1.55, w: 4.0, h: 0.3,
 fontSize: 12,
 fontFace: "Courier New",
 color: "E0E7FF"
 });
 slide.addText(" email=email,", {
 x: codeX + 0.3, y: codeY + 1.85, w: 4.0, h: 0.3,
 fontSize: 12,
 fontFace: "Courier New",
 color: "E0E7FF"
 });
 slide.addText(" )", {
 x: codeX + 0.3, y: codeY + 2.15, w: 4.0, h: 0.3,
 fontSize: 12,
 fontFace: "Courier New",
 color: "E0E7FF"
 });
 slide.addText(" db.session.add(user)", {
 x: codeX + 0.3, y: codeY + 2.45, w: 4.0, h: 0.3,
 fontSize: 12,
 fontFace: "Courier New",
 color: "E0E7FF"
 });
 slide.addText(" return user", {
 x: codeX + 0.3, y: codeY + 2.75, w: 4.0, h: 0.3,
 fontSize: 12,
 fontFace: "Courier New",
 color: "E0E7FF"
 });

 // 标签: AI 生成
 slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
 x: codeX + codeW - 1.0, y: codeY + 0.42, w: 0.8, h: 0.3,
 fill: { color: theme.accent },
 line: { type: 'none' },
 rectRadius: 0.05
 });
 slide.addText("AI", {
 x: codeX + codeW - 1.0, y: codeY + 0.42, w: 0.8, h: 0.3,
 fontSize: 10,
 fontFace: "Arial",
 color: "FFFFFF",
 bold: true,
 align: "center",
 valign: "middle"
 });

 // 底部记忆点
 slide.addShape(pres.shapes.RECTANGLE, {
 x: 0.5, y: 5.3, w: 9.0, h: 0.04,
 fill: { color: theme.accent },
 line: { type: 'none' }
 });
 slide.addText("亮点 / 从「逐字补全」升级为「按意图生成」, 并严格遵循项目现有代码风格", {
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
 pres.writeFile({ fileName: "slide-05-preview.pptx" });
}

module.exports = { createSlide, slideConfig };
