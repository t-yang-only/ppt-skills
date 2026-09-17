# taskpkg-data-mining-and-web-curation: 项目数据扫描萃取与网络高精配图审美质检规范

> 核心定位：
> 1. 深度扫描已有工程目录与项目文档，自动萃取核心性能指标、对比矩阵、演进时间线与架构清单并格式化输出；
> 2. 提供网络商业级配图智能检索，并对图片执行清晰度（拉普拉斯方差）、动态对比度、尺寸畸变等客观画质质检，确保合格后才装配入 PPT。

---

## 1. 项目数据扫描与可用资产萃取 (Project Data Harvesting)

脚本工具：`scripts/project_data_harvester.py`

### 核心功能与输出维度
1. **代码库与技术栈体检**：自动扫描项目各类代码与配置文件，汇总主要开发语言与核心依赖。
2. **核心 KPI 关键数据挖掘**：自动从 Markdown / TXT / JSON 文档中正则识别吞吐量（QPS/TPS）、延迟（ms）、百分比（%）、倍数等核心数字，封装为 4 组大字报 Hero 指标卡。
3. **架构层级与组件清单**：提炼接入层、中台业务层、模型算法层与数据底座清单。
4. **结构化对比矩阵 (Comparison Table)**：
   - 提取“传统方案 vs 本项目方案”多维对比数据；
   - 支持一键直接调用 python-pptx 将其生成为原生高审美表格插入目标幻灯片。

### CLI 快速运行
```bash
# 扫描工程并输出结构化 PPT 数据 JSON
python scripts/project_data_harvester.py "D:/project/my-repo" -o facts.json

# 扫描工程并将对比表直接插入指定 PPTX 第 3 页
python scripts/project_data_harvester.py "D:/project/my-repo" --pptx deck.pptx --slide 3
```

---

## 2. 网络配图检索与画质审美多维质检 (Web Image Curation & Aesthetic QA)

脚本工具：`scripts/web_image_curator.py`

为防止插入网络上常见的劣质模糊、尺寸畸变、带大面积水印或对比度灰暗的废片，建立四道硬性视觉质检关卡：

### 四道硬性视觉质检指标
1. **基础分辨率门禁**：宽度 ≥ 800px 且高度 ≥ 500px，杜绝缩略图与低保真模糊图。
2. **长宽比畸变检测**：$\max(w/h, h/w) \le 3.2$，防止插入极端狭长的广告横幅或碎片图。
3. **OpenCV 拉普拉斯方差清晰度检测 (Laplacian Variance)**：
   - 灰度图运行 $\sigma^2(\nabla^2 I) \ge 70.0$；
   - 自动识别失焦、重影与模糊废片，未达标直接淘汰拦截。
4. **RMS 动态对比度检测 (Contrast Dynamic Range)**：
   - 灰度标准差 $stddev \ge 28.0$；
   - 过滤曝光过度、全黑全白或泛白发灰的弱反差图片。

### CLI 快速运行
```bash
# 1. 搜索商业级科技配图，自动完成质检并下载前 3 张合格图
python scripts/web_image_curator.py --query "enterprise cloud architecture" --count 3 -o ./downloads

# 2. 对已有网络图片执行审美与画质体检
python scripts/web_image_curator.py --verify sample.jpg

# 3. 搜索配图、质检合格后直接裁剪为有机流体异形并插入 PPTX 第 2 页
python scripts/web_image_curator.py --query "artificial intelligence brain" --pptx deck.pptx --slide 2 --shape blob
```
