# 运行与验收依赖

## 必需依赖

- Python 3.10 或更高版本。
- `requirements.txt` 中的 `python-docx`，用于生成和检查 DOCX。

## 可选验收能力

- Microsoft Word 或 WPS 文字：面向教师交付时的目标编辑器验收。是否可直接控制取决于当前运行环境；不可用时必须报告未完成目标编辑器验收。
- LibreOffice 与 Poppler（`pdftoppm`）：仅供 `verify_docx.py` 做无头降级预览，二者不是生成 DOCX 的必需依赖，也不能替代 Word/WPS 验收。
- `fc-list`：若可用，`font_preflight.py` 用它列出字体；不可用时脚本会降级提示，不应据此自动替换目标端字体。

不要把本机应用路径、系统字体安装状态或某一个学校的版式要求视为所有使用者的默认条件。
