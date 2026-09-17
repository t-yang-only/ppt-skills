# taskpkg-vision-prompt-and-visual-editing: 视觉生图提示词工程与多维可视化编辑规范

> 核心定位：提供基于参考图反推完美生图提示词的工业级工程范式，并结合本地交互画板、Univer Slide 与 GenOffice 打造多维可视化 PPT 检视编辑工作流。

---

## 1. 完美生图提示词工程 (Vision Prompt Craft)

脚本工具：`scripts/image_prompt_craft.py`

针对 PPT 演示文稿所需的高精度、无干扰视觉素材，建立 **7 维结构化反推与提示词构筑范式**：

1. **核心主体 (Core Subject)**：精确描述核心物理形态、几何拓扑、机械装配或逻辑概念。
2. **艺术媒介与渲染器 (Artistic Medium & Engine)**：
   - 3D 磨砂玻璃拟态 (`3d-glass`)：半透明亚克力、霓虹微发光、高折射率、精细倒角。
   - 极简黏土低多边形 (`clay-minimal`)：柔和磨砂黏土质感、可爱比例、Blender 柔和棚拍。
   - 工业科技金属 (`metallic-tech`)：拉丝钛合金、倒角高光、微观机械组件、冷光指示。
   - 商业杂志扁平 (`editorial-flat`)：包豪斯极简矢量线条、精致剪影。
   - 商业静物微距 (`macro-product`)：哈苏中画幅相机特写、浅景深。
3. **材质物理属性 (Materials)**：次表面散射 (SSS)、环境光遮蔽 (AO)、光线追踪真实反射。
4. **影棚光路 (Lighting)**：轮廓光 (Rim Light)、柔光箱侧逆光、泛光漫反射。
5. **构图轴测 (Framing)**：等轴测 30° (Isometric)、英雄低仰角、留白黄金三分法。
6. **色彩协调 (Color Palette)**：自动通过图像量化算法提取主色与点缀色，确保与 PPT 主题色系严密呼应。
7. **背景隔离控制 (Background Isolation - 核心绝技)**：
   - 激活 `--transparent-asset` 模式，强制加入：`isolated on a clean seamless solid pure white background, studio product cutout, no shadows`
   - **巨大收益**：生成的图片主体完全独立在纯白底上，可被 `matting_cutout.py` 100% 完美一键抠出无噪点透明 PNG！

---

## 2. 三维可视化编辑与交互检视体系 (Visual Editing Ecosystem)

PPT-Skills 支持三种层级的可视化交互检视与编辑模式：

### 模式 A：本地轻量 Web 画板工作台 (`visual_editor.py`)
- **零外部依赖**：基于标准库启动 `python scripts/visual_editor.py deck.pptx --port 8765`。
- **视觉辅助线系统**：
  - 动态切换 **12 列网格参考标尺**，直观核查卡片对齐与吸附情况；
  - 动态开启 **5% 安全外边距红线框**，秒级排查边界压迫与溢出风险。
- **画板平移与放大**：支持多级滚轮缩放与像素级细节走查。

### 模式 B：平台级 Univer Slide 交互画布 (DSH Univer Suite)
- DSH 内置 Univer 运行时，支持将 PPTX 一键导入为交互式矢量幻灯片单元：
  - 调用 `univer_import` 将 `.pptx` 转为 `.univer`；
  - 调用 `sidebar_open` 在 DSH 侧边栏直接打开交互式多页画布；
  - 在侧边栏内可视化移动形状、修改文本、微调位置，并导出回 `.pptx`。

### 模式 C：桌面级 GenOffice 原生可视化编辑 (GenOffice Desktop)
- 调用 `genoffice open presentation.pptx`，直接调起本地 GenOffice 桌面可视化编辑窗口；
- 支持所见即所得 (WYSIWYG) 的文本微调、元素排布与格式重排。
