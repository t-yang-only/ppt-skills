// slide-03-pain-points.js - 开发者痛点场景
const pptxgen = require("pptxgenjs");

const FONT = "Microsoft YaHei";

const slideConfig = {
 type: 'content',
 index: 3,
 title: '开发者痛点'
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
 slide.addText("01 痛点场景", {
 x: 0.85, y: 0.3, w: 4.0, h: 0.3,
 fontSize: 11,
 fontFace: "Arial",
 color: theme.accent,
 bold: true,
 charSpacing: 3
 });

 // 主标题
 slide.addText("我们每天到底在为什么浪费时间?", {
 x: 0.5, y: 0.7, w: 9.0, h: 0.7,
 fontSize: 30,
 fontFace: FONT,
 color: theme.primary,
 bold: true
 });

 // 副标题
 slide.addText("作为开发者, 这些场景你一定不陌生: ", {
 x: 0.5, y: 1.4, w: 9.0, h: 0.4,
 fontSize: 14,
 fontFace: FONT,
 color: theme.secondary
 });

 // 4 个痛点卡片 (2x2 网格)
 const pains = [
 {
 icon: "1",
 title: "上下文切换成本高",
 desc: "在 Stack Overflow, 文档, IDE 之间反复跳转\n一个简单问题往往要查 10 分钟"
 },
 {
 icon: "2",
 title: "样板代码又长又啰嗦",
 desc: "每次新功能都要重写 CRUD, DTO, Controller\n时间花在重复劳动上"
 },
 {
 icon: "3",
 title: "Bug 定位耗时长",
 desc: "错误堆栈几千行, 不知道从哪下手\n修一个 bug 引入三个新 bug"
 },
 {
 icon: "4",
 title: "新代码不敢轻碰",
 desc: "没有测试, 文档, 注释\n重构像在雷区里走钢丝"
 }
 ];

 const cardW = 4.35;
 const cardH = 1.55;
 const startX = 0.5;
 const startY = 1.95;
 const gapX = 0.3;
 const gapY = 0.2;

 pains.forEach((p, i) => {
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

 // 左侧色条
 slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
 x: x, y: y, w: 0.1, h: cardH,
 fill: { color: theme.primary },
 line: { type: 'none' },
 rectRadius: 0.05
 });

 // 编号大圆
 slide.addShape(pres.shapes.OVAL, {
 x: x + 0.3, y: y + 0.3, w: 0.6, h: 0.6,
 fill: { color: theme.primary },
 line: { type: 'none' }
 });
 slide.addText(p.icon, {
 x: x + 0.3, y: y + 0.3, w: 0.6, h: 0.6,
 fontSize: 22,
 fontFace: "Arial",
 color: "FFFFFF",
 bold: true,
 align: "center",
 valign: "middle"
 });

 // 标题
 slide.addText(p.title, {
 x: x + 1.05, y: y + 0.3, w: cardW - 1.2, h: 0.4,
 fontSize: 16,
 fontFace: FONT,
 color: theme.primary,
 bold: true,
 valign: "middle"
 });

 // 描述
 slide.addText(p.desc, {
 x: x + 1.05, y: y + 0.75, w: cardW - 1.2, h: 0.7,
 fontSize: 11,
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
 slide.addText("事实 / 据行业调研, 开发者平均每天有 30% 的时间花在这些「非创造性」工作上", {
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
 pres.writeFile({ fileName: "slide-03-preview.pptx" });
}

module.exports = { createSlide, slideConfig };
