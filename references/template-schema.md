# 模板元数据 Schema

`parse_template.py` 输出的 `template_meta.json` 结构说明。

## 顶层字段

| 字段 | 类型 | 说明 |
|------|------|------|
| `source` | string | 模板文件绝对路径 |
| `global` | object | 全局信息（slide 尺寸） |
| `theme` | object | 主题字体与配色方案 |
| `layouts` | array | 所有母版版式 |
| `sample_slides` | array | 模板中已存在的示例页（如有） |
| `stats` | object | 字体/字号/颜色统计 |

## global

```json
{
  "slide_width_pt": 720.0,
  "slide_height_pt": 405.0
}
```
- 标准 16:9 模板：720 × 405 pt
- 标准 4:3 模板：720 × 540 pt

## theme

```json
{
  "major_font": "Calibri Light",
  "minor_font": "Calibri",
  "color_map": { "bg1": "lt1", "tx1": "dk1", ... }
}
```
- `major_font` / `minor_font` 来自 OOXML 的 `<a:majorFont>` / `<a:minorFont>`
- 字体方案未在母版中显式定义时，字段为 `null`

## layouts[]

每个版式结构：

```json
{
  "name": "Title and Content",
  "type": "cover" | "toc" | "section_header" | "content" | "summary" | "other",
  "placeholders": [...],
  "decorative_shapes": [...]
}
```

### placeholder 结构

```json
{
  "idx": 0,
  "type": "TITLE (1)" | "BODY (2)" | ...,
  "name": "Title 1",
  "position": {
    "left_pt": 60.0, "top_pt": 50.0,
    "width_pt": 600.0, "height_pt": 80.0
  },
  "sample_text": "示例文本" | null,
  "font": { "name": "Calibri", "size_pt": 36.0, "bold": true, "italic": false } | null
}
```

### decorative_shapes 结构

```json
{
  "name": "Rectangle 3",
  "type": "RECTANGLE (1)" | "PICTURE (13)" | ...,
  "position": { "left_pt": ..., "top_pt": ..., "width_pt": ..., "height_pt": ... },
  "is_picture": true,  // 仅图片时存在
  "fill_color": "#1F4E79"  // 仅纯色填充时存在
}
```

## stats

```json
{
  "fonts_top5": [["Calibri", 12], ["Calibri Light", 8], ...],
  "sizes_top10": [[36, 5], [18, 4], [14, 3], ...],
  "colors_top10": [["#1F4E79", 3], ["#FFFFFF", 2], ...]
}
```

## 用途

- **AI 排版决策**：根据 `theme.major_font/minor_font` 选主字体；根据 `stats.sizes_top10` 判断常见字号
- **版式选择**：根据 `layouts[].type` 选封面/目录/正文/总结版式
- **装饰还原**：`decorative_shapes` 中的 logo/色块/线条位置已保留，AI 不需在 content 中重复指定
