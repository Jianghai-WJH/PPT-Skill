# Slot 判定规则手册

本技能的核心创新：将 PPT 中的每个文字 run **逐个分类**为"模板保留"（template）或"变量替换"（variable）。

## 一、判定粒度

| 粒度 | 含义 | 适用 |
|------|------|------|
| **run 级** | 每个 `<a:r>` 单独判定 | ✅ 默认 |
| 段落级 | 整段统一判定 | 装饰段落 |
| 形状级 | 整个 shape 保留/删除 | 图片、纯装饰形状 |

## 二、判定规则（按优先级）

### 1. 模板（template）— 保留不替换

| 规则 | 例子 | 判定函数 |
|------|------|---------|
| 章节号（多级编号） | "3"、"3.3"、"3.3.1" | `is_chapter_like` |
| 中文/英文编号 | "1、"、"（1）" | `is_label_like` |
| 纯标点符号 | "："、"、"、"；" | `is_punctuation_only` |
| 固定模板标签词 | "5W2H"、"When"、"Where"、"Who"、"What"、"Why"、"How"、"How many" | `is_template_label_word` |
| 模板小标题词 | "故障现象"、"原因分析"、"临时对策"、"长久对策"、"对策实施计划" | `looks_like_template_heading` |
| 空 run | "" | （保留格式需要） |

### 2. 变量（variable）— 可被用户内容替换

| 规则 | 例子 | 判定函数 |
|------|------|---------|
| 纯数字 | "20"、"5"、"3.14" | `is_data_like` + `PURE_NUMBER_PATTERN` |
| 数字+单位 | "8L/min"、"15min"、"6℃" | `is_data_like` + `DATA_LIKE_PATTERN` |
| 描述性长句 | "轮毂电机车水温上升速度较基础车更慢" | 默认（启发式未命中 → variable） |
| 数据型日期 | "2025-10-09" | `is_data_like` |
| 计数 | "21#"、"3次" | `is_data_like` |
| 副标题/正文 | "整车除霜除雾性能较基础车变差" | 默认 |

### 3. 形状级（特殊处理）

| Shape 类型 | 默认动作 | 理由 |
|----------|---------|------|
| `placeholder_title` | edit_runs | 保留占位符，逐 run 判定 |
| `placeholder_body` | edit_runs | 同上 |
| `image` | delete | 默认删除（用户可重新插入） |
| `decoration` | delete | 装饰元素 |
| `logo` | keep | logo 保留 |
| `text_label` | edit_runs | 文本框/标签 |

## 三、组合（GROUP）递归

PPT 中常用"组合"组织多个子形状。slot 检测**必须递归**处理组合内每个子 shape，否则会丢失内容。

**关键**：每个 variable run 的 path 必须包含 sub_shape 路径：

```
sh2.sub0.p1.r2  # sh2 是个组合，进入 sub0 后的 p1 r2
sh2.sub1.p3.r0  # sh2 组合的另一个子形状
```

## 四、AI-Hint 模式

heuristic 模式的判定**不是 100% 准确**。对于复杂模板，建议用 `--mode ai-hint` 让 AI 做最后决策：

```bash
python detect_template_slots.py template.pptx slot_map.json --mode ai-hint
```

ai-hint 模式输出每个 run 的 `text + size + bold + heuristic role + 建议规则`，AI 看到这些后可以批量确认或调整。

## 五、典型场景的判定

### 场景 1：5W2H 模板

```
5W2H:        ← template (固定标签词)
When: 2025-10-09   ← template "When:" + variable "2025-10-09"
Where: 实验室      ← template "Where" + template ":" + variable "实验室"
Who: 试验员        ← template + variable
```

**用户替换时**：只填日期/地点/人，标签保持原样。

### 场景 2：编号列表

```
原因分析:        ← template (小标题)
1、原因1描述     ← template "1" + template "、" + variable "原因1描述"
2、原因2描述     ← 同上
```

**用户替换时**：保持编号"1、2"，只换描述。

### 场景 3：数据点

```
电机功率 7.4~8.7kW
              ↑       ↑
              var     var (含单位)
```

**用户替换时**：数字和单位都换。

## 六、扩展规则

要支持新的模板风格，可以修改 `detect_template_slots.py` 中的：

- `LABEL_PATTERNS`：新增编号格式（如 "①"、"A."）
- `FIXED_LABELS`：新增固定标签词
- `TEMPLATE_HEADING_HINTS`：新增模板小标题词
- `DATA_LIKE_PATTERN`：扩展单位列表
