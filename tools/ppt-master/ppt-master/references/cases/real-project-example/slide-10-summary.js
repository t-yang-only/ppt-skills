// slide-10-summary.js - 总结 + CTA
const pptxgen = require("pptxgenjs");

const FONT = "Microsoft YaHei";

const slideConfig = {
 type: 'summary',
 index: 10,
 title: '开始使用'
};

function createSlide(pres, theme) {
 const slide = pres.addSlide();
 slide.background = { color: theme.bg };

 // 大色块背景 (与封面呼应)
 slide.addShape(pres.shapes.RECTANGLE, {
 x: 0, y: 0, w: 10, h: 5.625,
 fill: { color: theme.primary },
 line: { type: 'none' }
 });
 slide.addShape(pres.shapes.RECTANGLE, {
 x: 0, y: 0, w: 10, h: 5.625,
 fill: { color: theme.secondary, transparency: 70 },
 line: { type: 'none' }
 });

 // 装饰大圆
 slide.addShape(pres.shapes.OVAL, {
 x: -2.0, y: -2.0, w: 5.0, h: 5.0,
 fill: { color: theme.accent, transparency: 50 },
 line: { type: 'none' }
 });
 slide.addShape(pres.shapes.OVAL, {
 x: 7.0, y: 3.0, w: 4.5, h: 4.5,
 fill: { color: theme.accent, transparency: 60 },
 line: { type: 'none' }
 });

 // 顶部小标签
 slide.addShape(pres.shapes.RECTANGLE, {
 x: 0.7, y: 0.7, w: 0.3, h: 0.04,
 fill: { color: theme.accent },
 line: { type: 'none' }
 });
 slide.addText("CLOSING", {
 x: 1.05, y: 0.6, w: 3.0, h: 0.3,
 fontSize: 11,
 fontFace: "Arial",
 color: theme.accent,
 bold: true,
 charSpacing: 4
 });

 // 大标题
 slide.addText("现在就试试", {
 x: 0.7, y: 1.0, w: 9.0, h: 0.9,
 fontSize: 56,
 fontFace: FONT,
 color: "FFFFFF",
 bold: true
 });
 slide.addText("MiniMax Code", {
 x: 0.7, y: 1.85, w: 9.0, h: 0.9,
 fontSize: 56,
 fontFace: "Arial",
 color: theme.accent,
 bold: true
 });

 // 副标
 slide.addText("三步开始: 装插件 -> 选项目 -> 说一句需求. 让 AI 帮你把活干完. ", {
 x: 0.7, y: 2.85, w: 9.0, h: 0.4,
 fontSize: 14,
 fontFace: FONT,
 color: theme.light
 });

 // 3 步卡片
 const steps = [
 { num: "01", title: "安装", desc: "VSCode / JetBrains\n插件市场一键安装" },
 { num: "02", title: "登录", desc: "MiniMax 账号\n免费个人版立即可用" },
 { num: "03", title: "开始", desc: "打开项目\n直接和 AI 协作" }
 ];

 const sW = 2.8;
 const sStartX = 0.7;
 const sY = 3.55;
 const sGap = 0.25;
 const sH = 1.05;

 steps.forEach((s, i) => {
 const x = sStartX + i * (sW + sGap);
 // 卡片
 slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
 x: x, y: sY, w: sW, h: sH,
 fill: { color: "FFFFFF", transparency: 88 },
 line: { color: "FFFFFF", width: 1, transparency: 70 },
 rectRadius: 0.1
 });
 // 编号
 slide.addText(s.num, {
 x: x + 0.2, y: sY + 0.1, w: 0.7, h: 0.4,
 fontSize: 22,
 fontFace: "Arial",
 color: theme.accent,
 bold: true
 });
 // 标题
 slide.addText(s.title, {
 x: x + 0.2, y: sY + 0.5, w: sW - 0.4, h: 0.3,
 fontSize: 16,
 fontFace: FONT,
 color: "FFFFFF",
 bold: true
 });
 // 描述
 slide.addText(s.desc, {
 x: x + 0.2, y: sY + 0.78, w: sW - 0.4, h: 0.3,
 fontSize: 9.5,
 fontFace: FONT,
 color: theme.light
 });
 });

 // CTA 链接
 slide.addText("minimax.io/code", {
 x: 0.7, y: 4.85, w: 5.0, h: 0.4,
 fontSize: 18,
 fontFace: "Arial",
 color: theme.accent,
 bold: true
 });
 slide.addText("hello@minimax.io", {
 x: 5.0, y: 4.85, w: 4.5, h: 0.4,
 fontSize: 14,
 fontFace: "Arial",
 color: "FFFFFF",
 align: "right"
 });

 // 底部 logo 文字
 slide.addText("MiniMax Code", {
 x: 0.7, y: 5.25, w: 5.0, h: 0.3,
 fontSize: 10,
 fontFace: "Arial",
 color: theme.light,
 charSpacing: 4
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
 pres.writeFile({ fileName: "slide-10-preview.pptx" });
}

module.exports = { createSlide, slideConfig };
