---
name: fig2ppt
description: 把论文框架图/流程图/架构图图片重建成原生可编辑形状的 PPTX（非贴图）。当用户说"图片/框架图/流程图/架构图做成可编辑 PPT/PPTX"、"把这张图转成能改的 PPT"、"图转矢量 PPT"时使用。核心流程：detect 出方框底稿 → AI 看图转写 spec JSON → build 生成原生形状 pptx → preview 渲染对比验收。
---

# fig2ppt：图片 → 原生可编辑 PPTX

把框架图重建成**可拖动、可改字**的原生 PPT 形状（不是把图片贴进 PPT）。

## 三条铁律（教训换来，违反必返工）

1. **绝不碰 PowerPoint COM 渲染**。曾连坏 4 次（Visible 参数、后台无输出、僵尸进程、拒绝调用）。核验只用 `scripts/preview_pptx.py`（matplotlib 渲染，唯一可靠通道）。
2. **detect 只出底稿是流程固有限制，不要试图全自动化**。detect 给"方框位置+颜色+形状猜测"，文字、箭头、层级必须靠 AI 看图转写 spec（裁图放大读字）。曾为此白迭代 5 轮算法。
3. **还原度"够用"即交付，用户说可以就立刻停**。验收标准：整体布局/配色/文字内容正确 = 可交付；不追求像素级一致。单图修正不超过 3 轮，仍有出入就在交付说明里标注遗留点并交付。

## 环境自检（每次启用先跑）

```bash
D:\Python\python.exe -c "import cv2, pptx, matplotlib, numpy, PIL; print('env ok')"
```

缺包装装：`D:\Python\python.exe -m pip install opencv-python python-pptx matplotlib numpy pillow`。
本机统一用 `D:\Python\python.exe`；换机器时换成该机的 python。

## 流程（每张图）

```bash
S=~/.qoder-cn/skills/fig2ppt/scripts
# 1. 底稿：识别方框位置/颜色/形状猜测 → <图名>.json + 调试叠加图
D:\Python\python.exe $S/detect.py <图.png> <图名>.json dbg_<图名>.png
# 2. 转写 spec：对照原图和 dbg 图写 spec_<图名>.json（字段见 reference/spec-format.md，
#    模板见 reference/spec-example.json）。文字/箭头/虚线容器全靠这一步人工判断。
# 3. 生成：spec 与 <图名>.json 放同一目录
D:\Python\python.exe $S/build_pptx.py spec_<图名>.json out.pptx
# 4. 核验：渲染预览，与原图并排对比，不合格回第 2 步改 spec（最多 3 轮，见铁律 3）
D:\Python\python.exe $S/preview_pptx.py out.pptx prev%d.png
```

多图合并成一个 pptx：把多个 spec 依次传给 build（`spec_1.json spec_2.json out.pptx`），一个 spec 一页。

## 关键契约

- spec 里 `"image": "xxx.png"` 决定 build 去 spec 同目录读 `xxx.json`（detect 产物）。**三者文件名必须对上**，放同一目录。
- detect 结果不理想时调参数重跑（`--min-area` 调小抓更碎的框、`--min-lum` 调低抓深色框、`--keep-caption` 保住被误裁的底部内容），**不改源码**。
- spec 写错 id 时 build 会打印 `WARN ... 不在 detect 结果中`，看到 WARN 先修 spec 再继续。

## 稳定性维护

- 改了 `scripts/build_pptx.py` 的 spec 字段行为，**必须同步更新** `reference/spec-format.md`。
- 改了任何脚本后，跑冒烟测试确认无回归：`D:\Python\python.exe ~/.qoder-cn/skills/fig2ppt/selftest/run_selftest.py`（约 30 秒，走通 detect→build→preview 全链路）。
- **脚本 docstring/注释只用 ASCII，别写中文路径**。打包时曾因 docstring 里出现 `技术路线图\_build` 触发 `SyntaxWarning: invalid escape sequence` + 终端 GBK 乱码（v1.0.0 已修，改脚本时别带回来）。
- 版本：v1.0.0（2026-09-26 自 技术路线图\_build 打包）。
