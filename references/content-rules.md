# content-rules.md（v3 多页版）

## 多页 content.json 完整示例

基于"AI云BMS-五看"模板的 3 页内容生成示例：

```json
{
  "slides": [
    {
      "use_slide_idx": 0,
      "fill_map": {
        "sh0.p0.r0": "清华大学-电驱安全组（汽车工程系）",
        "sh1.p0.r0": "电驱温度估算",
        "sh2.p0.r0": "核心痛点：温度传感器布点受限（通常 ≤4 路 NTC），但热模型需 30+ 节点。",
        "sh3.p0.r0": "Wang J. et al. Temperature estimation for EV powertrain. eTransportation 2025."
      },
      "clear_others": {
        "sh0.p0": 0,
        "sh1.p0": 0,
        "sh2.p0": 0,
        "sh3.p0": 0
      }
    },
    {
      "use_slide_idx": 1,
      "fill_map": {
        "sh0.p0.r0": "清华大学-电池管理团队（深圳国际研究生院）",
        "sh2.p0.r0": "用联邦学习取代单车训练，在云端聚合场景下，完成高精度 SOH 估计。",
        "sh3.p0.r0": "Li Z. et al. Federated learning for cross-vehicle SOH. Nat. Energy 2025."
      },
      "clear_others": {
        "sh0.p0": 0,
        "sh2.p0": 0,
        "sh3.p0": 0
      }
    },
    {
      "use_slide_idx": 2,
      "fill_map": {
        "sh0.p0.r0": "清华大学-碳管理组（深圳国际研究生院）",
        "sh1.p0.r0": "电驱系统全生命周期碳足迹",
        "sh2.p0.r0": "碳足迹数据要跟具体工艺路径绑定（直接回收 vs. 火法 vs. 湿法）。",
        "sh3.p0.r0": "Zhou G. et al. LCA for EV powertrain. Nat. Sustain. 2025."
      },
      "clear_others": {
        "sh0.p0": 0,
        "sh1.p0": 0,
        "sh2.p0": 0,
        "sh3.p0": 0
      }
    }
  ]
}
```

## 字段说明

### `slides` (数组)

每项对应一张输出 slide：

- **`use_slide_idx`** (整数): 引用源模板中的哪张 slide（0-based）
- **`fill_map`** (对象): 内容映射
  - 键: `sh{shape_idx}.p{para_idx}.r{run_idx}` 或 `sh{shape_idx}.sub{sub_idx}.p{p}.r{r}`
  - 值: 要填入的文本
- **`clear_others`** (对象, 可选): 指定段落清空其他 run
  - 键: `sh{shape_idx}.p{para_idx}`
  - 值: 保留的 run 索引（其他 run 的 text 会被设为空字符串）

### 多页生成的常用 pattern

**模式 1: 同一模板，N 张同结构页**
```json
{
  "slides": [
    {"use_slide_idx": 0, "fill_map": {...}},
    {"use_slide_idx": 0, "fill_map": {...}},
    {"use_slide_idx": 0, "fill_map": {...}}
  ]
}
```
适用：每页的"模板位置"完全一样，只是内容不同（如五看报告）

**模式 2: 跨模板页（封面 + 内容 + 封底）**
```json
{
  "slides": [
    {"use_slide_idx": 0, "fill_map": {...}},  // 引用封面版式
    {"use_slide_idx": 1, "fill_map": {...}},  // 引用内容版式
    {"use_slide_idx": 2, "fill_map": {...}}   // 引用总结版式
  ]
}
```
适用：综合 PPT（封面/目录/正文/总结混合）

**模式 3: 用户自定义 page structure**
```json
{
  "slides": [
    {"use_slide_idx": 0, "fill_map": {"sh0.p0.r0": "封面"}},
    {"use_slide_idx": 0, "fill_map": {"sh0.p0.r0": "目录"}},
    {"use_slide_idx": 0, "fill_map": {"sh0.p0.r0": "正文"}}
  ]
}
```
适用：每页都"克隆"源页+替换

## 排版规则

### 字号继承

不要在 fill_map 中指定字号——字号由**源模板的 run 样式**决定。新内容会**自动继承**原 run 的字号/字体/颜色/加粗。

### 颜色继承

- 普通文本：继承源 run 颜色（一般是黑色）
- 红色强调：源 run 是红色，新内容也会是红色（保持"红色加粗"格式）

### 段落 vs 整段替换

- `shX.pY.rZ` → 只替换**单个 run**（保留其他 run 不变）
- `clear_others: {shX.pY: 0}` + `shX.pY.r0` → **整段替换**（只保留 R0 样式，其他 run 文本清空）

## 模板与变量的判定

参考 [slot-judgment-rules.md](slot-judgment-rules.md)。

对于多页模板归纳，参考 [multi-page-template-extraction.md](multi-page-template-extraction.md)。
