// slide-08-usecases.js - 典型使用场景
const pptxgen = require("pptxgenjs");

const FONT = "Microsoft YaHei";

const slideConfig = {
 type: 'content',
 index: 8,
 title: '典型使用场景'
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
 slide.addText("04 典型用法", {
 x: 0.85, y: 0.3, w: 4.0, h: 0.3,
 fontSize: 11,
 fontFace: "Arial",
 color: theme.accent,
 bold: true,
 charSpacing: 3
 });

 // 主标题
 slide.addText("5 个高频场景", {
 x: 0.5, y: 0.7, w: 9.0, h: 0.7,
 fontSize: 30,
 fontFace: FONT,
 color: theme.primary,
 bold: true
 });
 slide.addText("覆盖开发者从「写新功能」到「修 bug」到「交付上线」的完整链路", {
 x: 0.5, y: 1.4, w: 9.0, h: 0.4,
 fontSize: 13,
 fontFace: FONT,
 color: theme.secondary
 });

 // 5 个时间轴节点 (横向流程)
 const scenarios = [
 {
 step: "01",
 title: "需求到代码",
 desc: "PRD / 一句话需求\n→ 直接生成可运行的功能模块"
 },
 {
 step: "02",
 title: "阅读陌生代码",
 desc: "快速理解一个陌生仓库\n生成结构图 + 阅读笔记"
 },
 {
 step: "03",
 title: "修 bug 找根因",
 desc: "贴入错误日志\n→ 5 分钟定位 + 修复方案"
 },
 {
 step: "04",
 title: "写测试 & 重构",
 desc: "补全测试覆盖率\n安全地做大规模重构"
 },
 {
 step: "05",
 title: "上线 & 文档",
 desc: "自动生成 PR 描述\nCHANGELOG, API 文档"
 }
 ];

 // 时间轴主线
 slide.addShape(pres.shapes.RECTANGLE, {
 x: 0.7, y: 2.85, w: 8.6, h: 0.04,
 fill: { color: theme.light },
 line: { type: 'none' }
 });

 const stepW = 1.6;
 const startX = 0.5;
 const gap = 0.2;
 const baseY = 2.7;

 scenarios.forEach((s, i) => {
 const x = startX + i * (stepW + gap);
 const cx = x + stepW / 2;

 // 圆点
 slide.addShape(pres.shapes.OVAL, {
 x: cx - 0.25, y: baseY, w: 0.5, h: 0.5,
 fill: { color: theme.primary },
 line: { color: "FFFFFF", width: 3 }
 });
 slide.addText(s.step, {
 x: cx - 0.25, y: baseY, w: 0.5, h: 0.5,
 fontSize: 11,
 fontFace: "Arial",
 color: "FFFFFF",
 bold: true,
 align: "center",
 valign: "middle"
 });

 // 上方标签
 slide.addText(s.title, {
 x: x, y: baseY - 0.5, w: stepW, h: 0.35,
 fontSize: 14,
 fontFace: FONT,
 color: theme.primary,
 bold: true,
 align: "center"
 });

 // 下方描述
 slide.addText(s.desc, {
 x: x, y: baseY + 0.65, w: stepW, h: 1.0,
 fontSize: 9.5,
 fontFace: FONT,
 color: theme.secondary,
 align: "center",
 valign: "top"
 });
 });

 // 底部"实战 demo"区
 slide.addShape(pres.shapes.ROUNDED_RECTANGLE, {
 x: 0.5, y: 4.7, w: 9.0, h: 0.55,
 fill: { color: theme.primary },
 line: { type: 'none' },
 rectRadius: 0.08
 });
 slide.addText("实战一句话 / ", {
 x: 0.7, y: 4.7, w: 1.4, h: 0.55,
 fontSize: 13,
 fontFace: FONT,
 color: theme.accent,
 bold: true,
 valign: "middle"
 });
 slide.addText("「给 Order 模块加一个按用户等级打折的接口, 并补齐测试」 -> MiniMax Code 5 分钟内给出 diff, 测试, PR 描述", {
 x: 2.0, y: 4.7, w: 7.4, h: 0.55,
 fontSize: 12,
 fontFace: FONT,
 color: "FFFFFF",
 valign: "middle"
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
 pres.writeFile({ fileName: "slide-08-preview.pptx" });
}

module.exports = { createSlide, slideConfig };
