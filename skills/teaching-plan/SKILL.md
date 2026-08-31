---
name: teaching-plan
description: 设计专业教案并生成 Word 文档。适用于根据课程要求创建教案、按已有 DOCX 参考版式重建教案，或提取可复用的教案模板；版式重建不等同于像素级复制，学校合规性须依据用户提供的制度核验。
---

# 教案设计与 Word 排版

根据课程要求生成结构化内容，再用本 Skill 内置脚本编译和验收 `.docx`。保留用户提供的事实；学校、教师、教材版本、课程标准和审查规则缺失时留空或向用户确认，不得虚构。

## 选择工作模式

- 用户提供参考 `.docx` 时，先运行 `scripts/inspect_docx.py`，再阅读 [样式规范](./references/style_schema.md) 与 [模板规范](./references/template_schema.md)，重建可观察到的页面、段落和表格结构。
- 用户未提供参考文档时，使用 `templates/default_vocational.json` 和 `styles/default_styles.json`。
- 用户只要求提取模板时，输出模板与样式 JSON，并说明未被提取的 DOCX 特性。

参考版式重建是基于脚本可提取属性的近似复现，不承诺像素级或 1:1 复制。页眉页脚、文本框、图片、域、内容控件、复杂样式继承、多节页面设置等未被完整提取时，必须明确报告。只有取得学校规范并逐项核验后，才可声称符合该校教务要求。

## 工作流

1. 收集课程、对象、课时、教材或课程标准、设备条件和交付要求。把信息分为已知事实、待确认事实和明确标注的临时假设。仅在会实质改变结果且无法从材料判断时追问；学校、教师、教材版本、课程标准、班级、设备和制度要求不得从示例或模板默认值补成用户事实。

   设备条件只来自当前请求。不得默认“有机房”或“无机房”；活动、替代方案和验收物必须适配用户实际提供的资源。
2. 参考版式模式运行：

   ```bash
   python <SKILL_PATH>/scripts/inspect_docx.py "<参考教案.docx>"
   ```

3. 阅读 [教学设计规范](./references/pedagogy_guidelines.md)，按 [内容 Schema](./references/content_schema.md) 生成内容 JSON。教学目标使用可观察行为；活动写明教师动作、学生操作、卡点与验收物，避免套话和虚构课堂细节。用户只要单课时或明确不要封面、进度表时，使用单课时内容范围，不得为满足完整模板而虚构整学期字段。
4. 验证结构。错误必须修复；警告必须核实：

   ```bash
   python <SKILL_PATH>/scripts/validate_content.py "<内容.json>"
   ```

   这里的 valid 仅表示内容结构可以进入渲染，不代表教学质量通过。

5. 按 [教学质量量表](./references/quality_rubric.md) 做质量评审。评审必须同时查看用户请求和成品，逐项记录证据：

   - 每个核心目标能定位到学生活动、验收物和判定标准；
   - 环节时间总和符合课时，任务量与课型相符；
   - 活动符合用户明确提供的设备、软件、网络和材料条件；
   - 困难学生有可执行支架，进阶学生有适量拓展；
   - 未知事实已追问、留空或标注适用条件，没有为了填满模板自行补全。

   命中量表阻断项时不得把教案标记为完成。维护本 Skill 或进行回归评测时，使用 evals/cases.json，并用 scripts/evaluate_quality.py 校验评分记录；正式回归还应为每个必评 case 传入 `--require-case`，并以 `--min-reviews 2` 验证双评审。评测夹具中的资源条件不是生成默认值。

6. 预检字体：

   ```bash
   python <SKILL_PATH>/scripts/font_preflight.py --style "<样式.json>"
   ```

   预检结果描述的是生成环境，不等于教师使用环境。宋体、黑体、楷体等目标端常用字体即使在生成环境缺失，也应默认保留原字体声明，并在 Word/WPS 中验收。只有用户要求当前机器也具备一致显示效果，或确认目标端同样缺少该字体时，才使用 `--strict` 将缺失视为阻断，并生成显式回退配置：

   ```bash
   python <SKILL_PATH>/scripts/font_preflight.py \
     --style "<样式.json>" \
     --write-override "<字体回退.json>"
   ```

7. 编译文档；存在字体回退配置时增加 `--override "<字体回退.json>"`：

   ```bash
   python <SKILL_PATH>/scripts/render.py \
     -c "<内容.json>" \
     -t "<模板路径或名称>" \
     -s "<样式路径或名称>" \
     -o "<输出.docx>"
   ```

8. 在目标编辑器中验收。优先使用 Computer Use 直接控制用户本机已安装的 Microsoft Word 或 WPS 文字打开输出文件；具体步骤见 [Word/WPS 验收](./references/editor_qa.md)。逐页检查缺字、裁切、断表、异常空白页、孤立反思页和页数异常。不要保存 Word/WPS 自动提出的格式升级或兼容性改写，除非用户明确要求。

9. 仅在 Word/WPS 不可用时，或需要无头环境中的快速结构回归时，运行 LibreOffice 降级预览：

   ```bash
   python <SKILL_PATH>/scripts/verify_docx.py \
     "<输出.docx>" \
     --output-dir "<新的QA目录>"
   ```

   该结果只能证明 LibreOffice 下的结构表现，不能替代 Word/WPS 验收，也不能据此改写原本面向 Word/WPS 的中文字体。若只能完成降级预览，交付时必须注明“尚未经过目标编辑器验收”。
10. 交付时提供文件路径、课程与学时摘要、结构验证与教学质量评审结果、目标编辑器及版本（若可见）、目标编辑器页数、字体状态、视觉检查页数，以及仍存在的版式差异或未核验项。任何交付门槛失败都不得把文件标记为完成。

## 运行约束

- 脚本路径均相对于本 Skill 目录解析，不硬编码外部 Skill 或工具安装路径。
- 渲染依赖见 `requirements.txt`；目标编辑器和降级预览的可选依赖见 [运行与验收依赖](./references/runtime_requirements.md)。缺少依赖时，使用当前项目的隔离 Python 环境从该文件安装；不要假定全局 `pip` 对应当前解释器。
- 不以“文件存在”或“能被 Word 打开”代替逐页视觉验收；也不以 LibreOffice 的页数和字体表现推断 Word/WPS 的最终效果。
