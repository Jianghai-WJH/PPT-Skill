# 版式类型说明

PPT 模板中常见的 5 类版式，本技能对每类都提供自动识别与对应填充逻辑。

## 1. cover（封面页）

- **用途**：文档首页，主标题 + 副标题 + 日期/作者
- **典型占位符**：TITLE（中心大标题）、SUBTITLE（副标题）、可能含日期占位符
- **识别关键词**：`Cover`, `Title Slide`, `封面`, `标题页`
- **content JSON 字段**：
  ```json
  { "layout_type": "cover", "title": "主标题", "subtitle": "副标题" }
  ```

## 2. toc（目录页）

- **用途**：展示文档大纲
- **典型占位符**：TITLE（"目录"/"Agenda"）、BODY（多行列表）
- **识别关键词**：`Table of Contents`, `Agenda`, `TOC`, `目录`
- **content JSON 字段**：
  ```json
  {
    "layout_type": "toc",
    "title": "目录",
    "items": ["项目背景", "技术路线", "方案设计", "风险评估"]
  }
  ```

## 3. section_header（章节首页）

- **用途**：分章节过渡，承上启下
- **典型占位符**：仅 TITLE
- **content JSON 字段**：
  ```json
  { "layout_type": "section_header", "title": "第三章 技术路线" }
  ```

## 4. content（正文页）

- **用途**：核心内容展示。最常用版式
- **典型占位符**：TITLE + BODY（支持项目符号）
- **识别关键词**：`Title and Content`, `Content`, `正文`, `标题与内容`
- **content JSON 字段**：
  ```json
  {
    "layout_type": "content",
    "title": "项目背景",
    "points": ["要点 1：xxx", "要点 2：xxx", "要点 3：xxx"],
    "table": {  // 可选
      "header": true,
      "rows": [
        ["指标", "现方案", "对标方案"],
        ["峰值功率", "200kW", "220kW"]
      ]
    },
    "notes": "演讲者备注"  // 可选
  }
  ```

## 5. summary（总结/封底）

- **用途**：关键结论、致谢、Q&A
- **典型占位符**：TITLE + 可能含 BODY
- **识别关键词**：`Thank You`, `谢谢`, `Summary`, `Conclusion`, `Q&A`
- **content JSON 字段**：
  ```json
  {
    "layout_type": "summary",
    "title": "总结",
    "points": ["核心结论 1", "核心结论 2", "下一步计划"]
  }
  ```

## 兜底策略

- 当模板中没有明确分类的版式时，`extract_layouts.py` 会对 `content` 类型回退到"占位符最多"的版式
- 当完全找不到目标版式时，`generate_pptx.py` 退回到第一个可用版式（通常 `Title Slide`）
