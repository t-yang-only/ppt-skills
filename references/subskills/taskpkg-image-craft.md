# PPT 视觉资产深度工坊与图像抠图排版手册 (Image Craft Pipeline)

本规范专为 PPT 高端视觉资产的工业化生产与装配而设计，涵盖「反向提示词生图 -> 区域切片 -> 透明扣图 -> 异形多边形裁剪 -> PPTX 原生精确定位装配」全流程。

---

## 1. 核心工作流与命令速查

所有视觉处理能力统一收拢在 `scripts/image_craft_pipeline.py` 中：

### 1.1 完美生图提示词工程 (`prompt`)
根据需求主体，自动生成包含 Octane/Blender 3D 渲染器、高阶光影与纯色分离背景的提示词：
```bash
python scripts/image_craft_pipeline.py prompt "智慧物流无人配送车芯片" --style 3d-render --palette deep-navy-neon-cyan --transparent
```

### 1.2 局部区域精确截取 (`crop`)
支持归一化坐标比例 (0.0~1.0) 截取画面核心主体：
```bash
python scripts/image_craft_pipeline.py crop raw_scene.png -o cropped_subject.png --bbox 0.15,0.20,0.85,0.80
```

### 1.3 高精无背景抠图与羽化 (`cutout`)
提取 Alpha 透明通道，自适应剔除背景色差并施加边缘抗锯齿微羽化：
```bash
python scripts/image_craft_pipeline.py cutout cropped_subject.png -o subject_transparent.png --tolerance 35
```

### 1.4 不规则形状裁剪与美化蒙版 (`mask`)
打破千篇一律的矩形框，将图片应用超采样抗锯齿异形蒙版：
- `organic-blob`: 呼吸感有机流体形态（高级杂志、未来科技首选）
- `hexagon`: 现代蜂窝六边形（工业互联网、架构图首选）
- `round-rect`: 苹果级大圆角卡片
- `circle`: 完美正圆形
- `snip-corner`: 极客斜切角硬核多边形

```bash
python scripts/image_craft_pipeline.py mask subject.png -o subject_blob.png --shape organic-blob
```

### 1.5 原生定位放置入 PPT (`place`)
将透明素材或异形图片以指定坐标 (英寸) 精确装配至 PPT 页面：
```bash
python scripts/image_craft_pipeline.py place presentation.pptx --slide 2 --image subject_blob.png --left 7.5 --top 2.2 --width 5.0
```

---

## 2. 视觉排版禁忌与防坑指南
1. **严禁整页大图强塞**：页面主体与文字必须分层解耦，图片尽量做透明抠图或异形蒙版处理；
2. **拒绝毛刺白边**：生图时务必附带 `--transparent` 启用纯色背景隔离，抠图后检查边缘羽化；
3. **呼吸感网格对齐**：异形图片放置时，外边界仍需尊重 12 列网格参考线，杜绝视觉重心失衡。
