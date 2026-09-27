# fig2ppt spec 格式速查（对应 scripts/build_pptx.py v1.0.0）

一个 spec JSON = 一页幻灯片。坐标单位 = **原图像素**（以 detect 产物 `<图名>.json` 的 W/H 为准），build 自动缩放居中到 13.33x7.5 英寸页面。

## 顶层字段

| 字段 | 类型 | 默认 | 说明 |
|---|---|---|---|
| `image` | string | **必填** | 原图文件名；build 据此在 spec 同目录找 `<去后缀>.json`（detect 产物） |
| `keep_only` | bool | false | true 时只保留 overrides 的键 + merges 的 ids 引用的框，其余全丢 |
| `drops` | [int] | [] | （keep_only=false 时）丢弃这些 detect id |
| `overrides` | {"id": {...}} | {} | 按 detect id 改框属性，见下表 |
| `merges` | [{ids,...}] | [] | 把多个 detect 框合并成一个大框 |
| `add` | [{rect,...}] | [] | 新增框（detect 没识别到的，如白底框、文字标签框） |
| `labels` | [{x,y,w,h,text,...}] | [] | 纯文字（无框文本框） |
| `arrows` | [...] | [] | 连线/箭头 |
| `polys` | [{pts,...}] | [] | 闭合多边形（异形色块） |
| `images` | [{rect,crop}] | [] | 从原图裁一块贴进去（实在重建不了的局部） |
| `default_shape` | string | "rect" | detect 框 shape 为空时的兜底形状 |
| `round` | bool | true | false 时 "rect" 用直角矩形而非圆角 |
| `font` | string | "Times New Roman" | 全图字体（中文图设 "SimHei"/"SimSun" 等） |

## overrides（键是 detect id 的字符串）

可改字段：`shape`、`fill`、`stroke`、`size`（字号 pt）、`dash`（虚线边）、`no_line`（无边）、`text`（框内文字，`\n` 换行）、`text_color`。

```json
"overrides": {"13": {"text": "Backbone", "shape": "triangleR", "fill": "#D8E8FF", "size": 11.05}}
```

## merges

```json
"merges": [{"ids": [3, 4, 5], "text": "Backbone", "shape": "rect", "fill": "#D8E8FF", "stroke": "#5B7DB1"}]
```
取 ids 的外接矩形；text/shape/fill/stroke 缺省继承第一个框。

## add（新增框）

必填 `rect: [x, y, w, h]`。可选：

| 字段 | 默认 | 说明 |
|---|---|---|
| `text` / `text_color` / `size` | "" / "#000000" / 自动 | 文字与字号；size 不写按框高宽自动换算 |
| `shape` | "rect" | 见形状表 |
| `fill` / `stroke` | "#FFFFFF" / "#000000" | 填充/描边色；stroke 省略时自动取 fill 的加深色；stroke="#FFFFFF" 会自动改深（防隐形） |
| `no_fill` / `no_line` | false | 无填充 / 无描边（配合 `dash:true` 画虚线容器） |
| `dash` | false | 虚线描边 |
| `bg` | false | 置底层（先画大容器再用它压底） |
| `icon` | - | `"stripes"`：框内画一叠平行四边形（特征图堆叠图标），`n` 控制层数（默认 4） |
| `rot` / `crot` | 0 | rot=绕中心旋转度数（宽高互换）；crot=以 (x,y) 为中心放置并旋转 |
| `adj` | - | 形状调节参数（圆角半径/六边形切角等，0~0.5） |
| `flipH` | false | 水平镜像 |
| `grad` | - | `{"colors": ["#A","#B"], "angle": 90}` 渐变填充 |
| `pattern` | - | `{"fg": "#A", "bg": "#B"}` 网格图案填充 |
| `stroke_w` | - | 描边宽度 pt |

## 形状表（shape 字段）

`rect`(圆角) `rect_sharp`(直角) `ellipse`/`ellipse_ish` `diamond` `triangleR/L/U`(三角，文字自动放框外) `hexagon`/`hexagon_v` `parallelogram` `trapezoid` `pentagon` `cylinder`(圆柱)

## arrows（连线）

两种写法：

```json
{"from": "12:r", "to": "13:l"}                          // 锚点：detect框id:边(l/r/t/b/c，缺省c)
{"from": "12:r", "to": "26:l", "mode": "vhv"}           // 肘形强制竖-横-竖（默认自动）
{"pts": [[69,155],[48,155],[48,362],[186,362]]}          // 手写折线点（最灵活，优先于锚点）
{"from": "12:r", "to": "13:l", "elbow": false}           // 直线（默认肘形）
{"from": "12:r", "to": "#300,200"}                       // 锚到自由坐标
```

公共可选：`color`(默认黑) `width`(pt, 默认1.0) `arrow`(是否有箭头，默认 true) `dash`(虚线，默认 false)。

## labels（纯文字）

```json
{"x": 240, "y": 114, "w": 117, "h": 26, "text": "Template vectors", "size": 16.57, "align": "l", "color": "#2E8B57", "bold": true}
```
`align`: c/l/r（默认 c）；size 不写自动。

## polys（闭合多边形）

```json
{"pts": [[x,y],...], "fill": "#A", "stroke": "#B", "width": 1.0, "dash": false, "bg": false}
```

## images（贴原图局部）

```json
{"rect": [x,y,w,h], "crop": [cx,cy,cw,ch]}   // crop 缺省 = rect；把原图 crop 区域缩放到 rect 位置
```
走投无路的局部（照片、复杂曲线图）才用，用了就不是全矢量。

## 行为要点（写 spec 前必知）

- **框按面积从大到小画**：大容器天然在底层；`bg:true` 是双保险。
- 三角形 triangleR/L 的文字不进形状，自动变成旁边文本框。
- 自动字号公式：`min(框高/(行数*1.28), 框宽/(0.52*最长行字符数)*1.05)`，钳制 5.5~20pt。不满意就用 `size` 显式指定。
- detect 形状猜测（rect/diamond/ellipse_ish/triangleX）只是初稿，不对就在 overrides 里改。
- 虚线容器标准画法：add + `no_fill:true, dash:true, stroke:"#色"`。
- build 打印 `WARN ... 不在 detect 结果中` = spec 引用了不存在的 id，先修再继续。
