# 多页模板归纳（v3 核心能力）

## 核心问题

用户的源 PPT 里有 **N 张同模板不同内容** 的 slide。技能要从这 N 张 slide 中**归纳出"什么是模板共性"**，然后生成新内容页。

## 工作流

```
源 PPT（N 张同模板 slide）
    ↓
extract_template_features.py
    ↓
template_features.json
{
  "num_slides_analyzed": N,
  "zones": {
    "header_institution": { ... 共性特征 ... },
    "title_main": { ... 共性特征 ... },
    "body_left": { ... 共性特征 ... },
    "footer_source": { ... 共性特征 ... }
  }
}
    ↓
AI 读取 features
    ↓
content.json（按 zone 填入新内容）
    ↓
generate_pptx.py
    ↓
最终 .pptx
```

## 共性归纳算法

### Step 1: 提取每张 slide 的特征

对每张 slide，提取：
- 每个 shape 的位置（x, y）、尺寸（w, h）
- 每个 run 的字号、字体、颜色、加粗
- 形状类型（text/image/table/group）

### Step 2: 按 zone 分组

定义 zone（相对位置划分）：
| Zone | 相对 y | 相对 x | 典型用途 |
|------|-------|-------|---------|
| `header_institution` | 0 - 0.12 | 任意 | 顶部机构标签 |
| `title_main` | 0.12 - 0.22 | 任意 | 主标题（22pt+ 加粗） |
| `body_left` | 0.22 - 0.85 | 0 - 0.55 | 左侧正文 |
| `body_right` | 0.22 - 0.85 | 0.55 - 1 | 右侧图表/图片 |
| `footer_source` | 0.85 - 1 | 任意 | 底部出处/法规 |

### Step 3: 统计共性

对每个 zone，统计：
- **出现次数**：3/3 slide 都出现 → 强共性
- **常见字号**：在 3 张 slide 中都出现的字号
- **常见颜色**：在 3 张 slide 中都出现的颜色
- **常见字体**：在 3 张 slide 中都出现的字体
- **位置范围**：x/y 坐标的 min-max

### Step 4: 推断变量规则

每个 zone 的"什么是变量"：
- 文本内容（如"清华-电驱安全组" vs "清华-电池安全实验室"）→ 变量
- 字号/颜色/字体/位置 → 模板（保留）

## 典型输出

```json
{
  "num_slides_analyzed": 3,
  "zones": {
    "header_institution": {
      "role": "text",
      "occurrences": 3,
      "common_size": 16.0,
      "common_bold": true,
      "common_color": "FF0000",
      "common_font": "微软雅黑",
      "x_range": [11, 11],
      "y_range": [13, 13],
      "variable_rules": {
        "text_is_variable": true,
        "format_is_template": true,
        "emphasis_pattern": "title_with_red_emphasis"
      }
    },
    "title_main": {
      "role": "text",
      "occurrences": 3,
      "common_size": 22.0,
      "common_bold": true,
      "common_color": "FF0000",
      "common_font": "微软雅黑",
      "x_range": [0, 495],
      "y_range": [54, 111],
      "variable_rules": {
        "text_is_variable": true,
        "format_is_template": true,
        "emphasis_pattern": "main_title"
      }
    },
    "body_left": {
      "role": "text",
      "occurrences": 3,
      "common_size": 16.0,
      "common_bold": null,
      "common_color": "FF0000",
      "common_font": "微软雅黑",
      "x_range": [-0, 379],
      "y_range": [120, 435],
      "variable_rules": {
        "text_is_variable": true,
        "format_is_template": true,
        "emphasis_pattern": "body_text"
      }
    },
    "footer_source": {
      "role": "text",
      "occurrences": 3,
      "common_size": 12.0,
      "common_bold": true,
      "common_color": "000000",
      "common_font": "微软雅黑",
      "x_range": [-0, 607],
      "y_range": [464, 505],
      "variable_rules": {
        "text_is_variable": true,
        "format_is_template": true,
        "emphasis_pattern": "body_text"
      }
    }
  }
}
```

## AI 如何使用

拿到 `template_features.json` 后，AI 应当：

1. **理解模板结构**：每个 zone 是"该填什么类型的内容"
2. **决定变量**：哪些 zone 是变量（一般 4 个都是变量）
3. **构建 content.json**：每张新内容页 → 一组 fill_map + clear_others
4. **保留格式**：填入新文本时，**不要**在 content.json 中指定字号/颜色——让模板的 run 样式自动继承

### 模板特征到内容 JSON 的映射

| zone | 推荐 content 字段 | 示例 |
|------|-----------------|------|
| header_institution | `institution` | "清华大学-电驱安全组" |
| title_main | `title` | "电驱温度估算" |
| body_left | `body` | "核心痛点：..." |
| footer_source | `source` | "Wang J. et al. ..." |

## 鲁棒性处理

### 位置容差

- y 坐标容差：±10px
- 如果 y=54 和 y=64 都被分到不同 zone，按"特征相似度"合并

### 字号差异

- 主标题在 22pt 附近（20-24pt 都算）
- 左侧正文在 10-16pt 范围

### 模板混合

- 当一个 slide 同时是 22pt 标题 + 16pt 正文 → 都识别为"标准内容页"模板

## 验证

跑 `extract_template_features.py` 后，检查：
- `num_slides_analyzed == N`
- `zones` 字典非空
- 每个 zone 的 `occurrences >= 2`（强共性是 3/3）

如果不满足，检查源 PPT 是否真的同模板（不同模板的特征无法归纳）。
