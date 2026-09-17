// slide-06-capability-context.js - 能力 2: 项目级上下文理解
const pptxgen = require("pptxgenjs");

const FONT = "Microsoft YaHei";

const slideConfig = {
 type: 'content',
 index: 6,
 title: '能力二:项目级上下文理解'
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
 slide.addText("02", {
 x: 0.5, y: 0.6, w: 1.2, h: 0.9,
 fontSize: 64,
 fontFace: "Arial",
 color: theme.light,
 bold: true
 });

 // 主标题
 slide.addText("项目级上下文理解", {
 x: 1.6, y: 0.8, w: 7.0, h: 0.6,
 fontSize: 30,
 fontFace: FONT,
 color: theme.primary,
 bold: true
 });
 slide.addText("不只是看当前文件, 而是真的懂你的项目", {
 x: 1.6, y: 1.4, w: 7.5, h: 0.4,
 fontSize: 13,
 fontFace: FONT,
 color: theme.secondary
 });

 // 左侧: 可视化项目结构 (中心辐射图风格)
 const centerX = 2.7;
 const centerY = 3.5;

 // 中心圆
 slide.addShape(pres.shapes.OVAL, {
 x: centerX - 0.7, y: centerY - 0.5, w: 1.4, h: 1.0,
 fill: { color: theme.primary },
 line: { type: 'none' }
 });
 slide.addText("MiniMax\nCode", {
 x: centerX - 0.7, y: centerY - 0.5, w: 1.4, h: 1.0,
 fontSize: 14,
 fontFace: FONT,
 color: "FFFFFF",
 bold: true,
 align: "center",
 valign: "middle"
 });

 // 周围 6 个节点
 const nodes = [
 { label: "代码文件", x: centerX - 1.4, y: centerY - 1.5, color: theme.accent },
 { label: "依赖关系", x: centerX + 0.5, y: centerY - 1.4, color: theme.secondary },
 { label: "Git 历史", x: centerX + 1.4, y: centerY - 0.3, color: theme.primary },
 { label: "测试用例", x: centerX + 0.5, y: centerY + 0.6, color: theme.accent },
 { label: "配置文件", x: centerX - 1.4, y: centerY + 0.6, color: theme.secondary },
 { label: "API 文档", x: centerX - 1.7, y: centerY - 0.5, color: theme.primary }
 ];

 nodes.forEach((n) => {
 // 连接线
 slide.addShape(pres.shapes.LINE, {
 x: centerX, y: centerY, w: n.x + 0.4 - centerX, h: n.y + 0.25 - centerY,
 line: { color: theme.light, width: 1.5 }
 });
 // 节点
 slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
 x: n.x, y: n.y, w: 0.85, h: 0.5,
 fill: { color: n.color },
 line: { type: 'none' },
 rectRadius: 0.08
 });
 slide.addText(n.label, {
 x: n.x, y: n.y, w: 0.85, h: 0.5,
 fontSize: 10,
 fontFace: FONT,
 color: "FFFFFF",
 bold: true,
 align: "center",
 valign: "middle"
 });
 });

 // 右侧: 4 个能力点
 const points = [
 {
 title: "全局检索",
 desc: "跨数千文件秒级定位, 理解函数调用链和模块边界"
 },
 {
 title: "约定遵循",
 desc: "学习项目的命名, 风格, 架构模式, 生成的代码与现有代码融为一体"
 },
 {
 title: "影响分析",
 desc: "修改前先告诉你「会动到哪些文件」「可能影响哪些测试」"
 },
 {
 title: "跨文件重构",
 desc: "一次命令完成跨多个文件的重命名, 提取, 抽象等重构"
 }
 ];

 const fX = 5.3;
 const fStartY = 1.95;
 const fH = 0.7;
 const fGap = 0.08;

 points.forEach((f, i) => {
 const y = fStartY + i * (fH + fGap);
 // 卡片
 slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
 x: fX, y: y, w: 4.2, h: fH,
 fill: { color: theme.light },
 line: { type: 'none' },
 rectRadius: 0.06
 });
 // 编号
 slide.addText(String(i + 1).padStart(2, "0"), {
 x: fX + 0.15, y: y + 0.05, w: 0.6, h: 0.4,
 fontSize: 16,
 fontFace: "Arial",
 color: theme.accent,
 bold: true
 });
 // 标题
 slide.addText(f.title, {
 x: fX + 0.8, y: y + 0.05, w: 3.3, h: 0.32,
 fontSize: 14,
 fontFace: FONT,
 color: theme.primary,
 bold: true
 });
 // 描述
 slide.addText(f.desc, {
 x: fX + 0.8, y: y + 0.36, w: 3.3, h: 0.35,
 fontSize: 9.5,
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
 slide.addText("亮点 / 上下文不是「塞进 prompt」, 而是基于项目结构的真正理解", {
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
 pres.writeFile({ fileName: "slide-06-preview.pptx" });
}

module.exports = { createSlide, slideConfig };
