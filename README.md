# fig2ppt

把论文框架图 / 流程图 / 架构图**图片**重建成**原生可编辑形状**的 PPTX —— 不是把图片贴进 PPT，生成的每个方框、文字、箭头都可以在 PowerPoint 里直接拖动、改字、换色。

## 工作流程

```
图片 ──detect──▶ 方框底稿(位置/颜色/形状) ──AI 看图转写──▶ spec JSON ──build──▶ 可编辑 PPTX ──preview──▶ 渲染图对比验收
```

1. **detect**：OpenCV 识别图中填充方框的位置、颜色、形状猜测，输出 JSON + 调试叠加图
2. **转写 spec**：AI 对照原图人工判断文字、箭头、虚线容器、层级，写成 spec JSON（这是半自动流程的关键一步，全自动不可靠）
3. **build**：spec + detect 底稿合成原生形状的 .pptx（python-pptx）
4. **preview**：matplotlib 把 pptx 渲染成 PNG，与原图并排对比验收（不依赖 PowerPoint）

## 目录结构

```
fig2ppt/                  # skill 本体
├── SKILL.md              # Qoder skill 定义（流程、铁律、维护规范）
├── scripts/
│   ├── detect.py         # 图片 → 方框底稿 JSON + 调试叠加图
│   ├── build_pptx.py     # spec JSON → 原生可编辑 PPTX
│   └── preview_pptx.py   # PPTX → PNG 预览（matplotlib 渲染）
├── reference/
│   ├── spec-format.md    # spec JSON 字段规范
│   └── spec-example.json # spec 模板示例
└── selftest/             # 冒烟测试：约 30 秒走通 detect→build→preview 全链路
    └── run_selftest.py
```

## 环境依赖

```bash
pip install opencv-python python-pptx matplotlib numpy pillow
```

## 快速开始

```bash
S=fig2ppt/scripts

# 1. 底稿：识别方框位置/颜色/形状 → <图名>.json + 调试叠加图
python $S/detect.py 图.png 图名.json dbg_图名.png

# 2. 转写 spec：对照原图和 dbg 图写 spec_图名.json
#    字段规范见 fig2ppt/reference/spec-format.md，模板见 spec-example.json

# 3. 生成：spec 与 detect 产物放同一目录
python $S/build_pptx.py spec_图名.json out.pptx

# 4. 核验：渲染预览，与原图并排对比
python $S/preview_pptx.py out.pptx prev%d.png
```

多图合并成一个 pptx（一个 spec 一页）：

```bash
python $S/build_pptx.py spec_1.json spec_2.json out.pptx
```

## 冒烟测试

改动任何脚本后跑一遍确认无回归：

```bash
python fig2ppt/selftest/run_selftest.py   # 输出 SELFTEST OK 即通过
```

## 设计要点

- **不碰 PowerPoint COM 自动化**：预览核验只用 matplotlib 渲染，跨机器稳定
- **半自动是刻意设计**：detect 只出几何底稿，文字/箭头/层级靠 AI 看图转写，避免脆弱的全自动识别
- spec 中 `"image": "xxx.png"` 决定 build 读取同目录的 `xxx.json`（detect 产物），三者文件名需对应

## 版本

v1.0.0（2026-09-26）
