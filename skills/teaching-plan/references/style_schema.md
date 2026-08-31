# 样式系统配置 (Styles JSON) 规范与参数字典

样式系统负责管理所有物理排版属性，包括页面边距、字体体系、字号层级、行间距以及表格几何尺寸。

---

## 1. 样式顶级结构

```json
{
  "page": { ... },            // 页面物理尺寸与页边距
  "fonts": { ... },           // 全局中西文字体映射
  "technical_text": { ... },  // 技术文本的西文字体
  "cover_style": { ... },     // 封面排版样式
  "schedule_style": { ... },  // 进度表排版样式
  "lesson_style": { ... }     // 课时大表单元格与行距排版样式
}
```

---

## 2. 核心属性字典

### 2.1 页面设置 (`page`)
*   `page_width_pt`: A4 纸宽 595.3 pt (21.0 cm)
*   `page_height_pt`: A4 纸高 841.9 pt (29.7 cm)
*   `top_margin_pt` / `bottom_margin_pt`: 上下边距（标准推荐 56.7 pt，即 2.0 cm）
*   `left_margin_pt` / `right_margin_pt`: 左右边距（标准推荐 84.6 pt，即 2.98 cm）

### 2.2 字体体系 (`fonts`)
*   `default_font`: 中文正文字体，默认 `"宋体"`
*   `ascii_font`: 西文/数字/代码字体，默认 `"Times New Roman"`
*   `heading_font`: 标题字体，默认 `"微软雅黑"` 或 `"黑体"`
*   `kaiti_font`: 强调/落款字体，默认 `"楷体_GB2312"`
*   `color_rgb`: 默认文字颜色，`"000000"`（纯黑）

字体名称只声明期望字体，不会把字体嵌入 DOCX，也不保证目标系统已安装。预检用于识别生成环境与目标环境的差异风险：

```bash
python <SKILL_PATH>/scripts/font_preflight.py --style "<样式.json>"
```

需要机器读取时增加 `--json`。只有当前生成机器也必须具备一致字体时才增加 `--strict`。结果中的 `suggested_fallback` 只是当前系统上已发现的候选；字体名称可被发现也不代表目标编辑器一定能正确渲染中文字形。若确认需要替换并接受候选字体，生成显式 override：

```bash
python <SKILL_PATH>/scripts/font_preflight.py \
  --style "<样式.json>" \
  --write-override "<字体回退.json>"
```

override 使用 `fonts.substitutions` 保存“原字体 → 已安装字体”映射，渲染时通过 `render.py --override "<字体回退.json>"` 应用。脚本不会静默改写原样式。若目标是教师电脑中的 Word/WPS，生成端缺少宋体、黑体或楷体不应自动触发替换；保留目标字体并在实际 Word/WPS 中验收。只有确认目标端同样缺失字体或用户要求跨平台一致显示时才应用 override。LibreOffice 的字体和分页结果仅作降级回归，不能代表 Word/WPS。

### 2.3 字号对照参考 (pt 与中文字号对应)
*   **初号**：42.0 pt (常用于封面主标题)
*   **小初**：36.0 pt
*   **一号**：26.0 pt
*   **二号**：22.0 pt
*   **小二**：18.0 pt (常用于进度表标题)
*   **三号**：16.0 pt (常用于封面副标题/落款)
*   **四号**：14.0 pt (常用于课时序号小标题)
*   **小四**：12.0 pt (常用于表格内正文与表头)
*   **五号**：10.5 pt

### 2.4 表格排版 (`cell_format`)
*   `line_spacing`: 行距倍数（如 `1.15` 倍）
*   `space_before_pt` / `space_after_pt`: 段前段后间距（如 `1.5 pt`）
*   `subhead_bold_keywords`: 自动识别并加粗的子标题关键字列表（如 `["【", "活动1", "教师活动", "学生活动"]`）

### 2.5 技术文本 (`technical_text`)

默认配置为 `{"enabled": true, "font_name": "Courier New"}`。在非整段加粗的表格正文中，保守识别常见 Linux 命令及选项、Unix 权限字符串、ASCII 路径、有效 IPv4 地址和 CIDR；只改变识别片段的西文字体，保留原字符、中文字体、字号及加粗属性。普通英文术语不一律视为代码，IPv6、未知命令及含中文路径不保证识别。

技术片段使用西文脚本提示，包含技术文本的段落同时关闭西文单词内逐字符换行及中西文自动间距，避免 Word 把 ASCII token 当作东亚文字逐字断开。该设置不是绝对禁止换行：斜杠、连字符等合法标点断点、过长路径以及窄列仍可能换行，必须在目标 Word/WPS 中检查权限串、参数、路径和 IP 是否容易误读。不要插入不可见字符或替换普通连字符来强行禁止断行，以免影响复制执行。需要严格代码展示时，应在内容设计阶段提供独立命令段，并检查可用列宽。

自定义等宽字体可覆盖 `technical_text.font_name`；关闭该功能使用 `technical_text.enabled: false`。字体不嵌入文档，目标端是否实际采用仍以编辑器为准。
