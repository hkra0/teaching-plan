# 教案内容范围与质量证据

## 内容范围和事实状态

默认是包含封面、进度表和课次的完整教案。用户只要求一个课时，或明确不要封面、进度表时，可使用以下顶层字段：

~~~json
{
  "document_profile": "single_lesson",
  "document_status": "final",
  "output_filename": "Excel绝对引用课堂教案.docx",
  "lessons": []
}
~~~

- document_profile 可为 full 或 single_lesson；未填写时按 full。
- document_status 可为 draft 或 final；未填写时按 final。
- single_lesson 不要求 cover 和 schedule，不得为满足完整模板虚构学校、教师或整学期数据。
- draft 中未知的学校、教师、教材和审核人可以留空并产生待确认警告。
- final 必须解决交付范围内的必填事实；无法确认时应保持草稿状态或缩小交付范围。

生成时区分已知事实、待确认事实和临时假设。假设必须说明适用条件，不能伪装成用户信息。

## 质量证据字段

下列字段不改变当前 Word 表格结构，但能为质量评审提供证据。新生成的质量就绪教案应优先提供：

- lesson.duration_minutes：本课可用分钟数。
- stage.minutes：本环节分钟数。
- stage.objective_refs：本环节服务的目标编号数组，例如 K1、A1。
- stage.evidence：学生留下的可检查产出或观察证据。
- stage.criterion：合格所需的必备内容、正确率、错误上限或检查步骤。
- stage.support：常见卡点触发时使用的分步支架或补救任务。
- stage.extension：提前完成或基础较好的学生使用的拓展任务。

目标文本可使用 K1、A1、Q1 等稳定编号。每个核心目标至少被一个活动引用，并在评价环节找到 evidence 和 criterion。

环节时间总和应等于 lesson.duration_minutes。实践比例依据课型和用户要求判断，不设跨学科固定默认值。

设备、软件、网络、账号和材料只能来自当前请求。任何示例中的资源条件都不能成为其他请求的默认条件。

