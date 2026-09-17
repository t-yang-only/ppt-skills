// slide-04-what-is.js - MiniMax Code 是什么
const pptxgen = require("pptxgenjs");

const FONT = "Microsoft YaHei";

const slideConfig = {
 type: 'content',
 index: 4,
 title: 'MiniMax Code 是什么'
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
 slide.addText("02 产品定义", {
 x: 0.85, y: 0.3, w: 4.0, h: 0.3,
 fontSize: 11,
 fontFace: "Arial",
 color: theme.accent,
 bold: true,
 charSpacing: 3
 });

 // 主标题
 slide.addText("不止是 Copilot", {
 x: 0.5, y: 0.7, w: 9.0, h: 0.7,
 fontSize: 30,
 fontFace: FONT,
 color: theme.primary,
 bold: true
 });
 slide.addText("而是一个真正懂项目的 AI 编程伙伴", {
 x: 0.5, y: 1.4, w: 9.0, h: 0.4,
 fontSize: 14,
 fontFace: FONT,
 color: theme.secondary
 });

 // 左侧大定义卡片
 slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
 x: 0.5, y: 1.95, w: 4.5, h: 3.0,
 fill: { color: theme.primary },
 line: { type: 'none' },
 rectRadius: 0.15
 });
 // 装饰圆
 slide.addShape(pres.shapes.OVAL, {
 x: 3.5, y: 2.0, w: 1.3, h: 1.3,
 fill: { color: theme.accent, transparency: 50 },
 line: { type: 'none' }
 });
 slide.addShape(pres.shapes.OVAL, {
 x: 0.3, y: 4.1, w: 1.5, h: 1.5,
 fill: { color: theme.secondary, transparency: 60 },
 line: { type: 'none' }
 });

 slide.addText("MiniMax Code", {
 x: 0.7, y: 2.15, w: 3.5, h: 0.4,
 fontSize: 14,
 fontFace: "Arial",
 color: theme.light,
 bold: true
 });
 slide.addText("AI Coding\nCompanion", {
 x: 0.7, y: 2.5, w: 3.5, h: 1.3,
 fontSize: 38,
 fontFace: "Arial",
 color: "FFFFFF",
 bold: true
 });
 slide.addShape(pres.shapes.RECTANGLE, {
 x: 0.7, y: 3.85, w: 0.5, h: 0.04,
 fill: { color: theme.accent },
 line: { type: 'none' }
 });
 slide.addText("深度集成于 IDE", {
 x: 0.7, y: 3.95, w: 3.5, h: 0.3,
 fontSize: 12,
 fontFace: FONT,
 color: theme.light
 });
 slide.addText("理解整个项目上下文 / 主动协同", {
 x: 0.7, y: 4.25, w: 3.5, h: 0.3,
 fontSize: 11,
 fontFace: FONT,
 color: "FFFFFF"
 });

 // 右侧 4 个特征点
 const features = [
 {
 title: "Agentic 工作流",
 desc: "不只是问答, 而是能拆解任务, 调用工具, 自主完成多步操作"
 },
 {
 title: "项目级上下文",
 desc: "理解整个代码库的依赖, 约定和架构, 而非孤立的代码片段"
 },
 {
 title: "可解释的决策",
 desc: "给出修改时说明原因, 修改有 diff, 可审查, 可回滚"
 },
 {
 title: "本地优先",
 desc: "敏感代码不出本地, 核心数据隐私可控, 企业可私有部署"
 }
 ];

 const fX = 5.3;
 const fStartY = 1.95;
 const fH = 0.72;
 const fGap = 0.05;

 features.forEach((f, i) => {
 const y = fStartY + i * (fH + fGap);

 // 圆点
 slide.addShape(pres.shapes.OVAL, {
 x: fX, y: y + 0.15, w: 0.2, h: 0.2,
 fill: { color: theme.accent },
 line: { type: 'none' }
 });

 // 标题
 slide.addText(f.title, {
 x: fX + 0.35, y: y, w: 4.2, h: 0.35,
 fontSize: 14,
 fontFace: FONT,
 color: theme.primary,
 bold: true
 });

 // 描述
 slide.addText(f.desc, {
 x: fX + 0.35, y: y + 0.35, w: 4.2, h: 0.4,
 fontSize: 10,
 fontFace: FONT,
 color: theme.secondary,
 valign: "top"
 });
 });

 // 底部记忆点
 slide.addShape(pres.shapes.RECTANGLE, {
 x: 0.5, y: 5.05, w: 9.0, h: 0.04,
 fill: { color: theme.accent },
 line: { type: 'none' }
 });
 slide.addText("一句话 / MiniMax Code = 懂项目的 AI 队友 + 可控的自动化 + 真正落地到 IDE", {
 x: 0.5, y: 5.15, w: 9.0, h: 0.3,
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
 pres.writeFile({ fileName: "slide-04-preview.pptx" });
}

module.exports = { createSlide, slideConfig };
