# 教学设计内容数据 (Content JSON) 结构规范

教案数据采用纯 JSON 格式描述，完全不包含排版杂质。

单课时范围、草稿未知字段和质量证据字段见 [内容范围与质量证据](./content_profiles.md)。

---

## 1. 完整 JSON 结构概览

```json
{
  "output_filename": "Python程序设计教学设计（示例）.docx",
  "template": "default_vocational",
  "cover": {
    "title": "教  学  设  计",
    "subtitle": "（示例学年 第一学期）",
    "info_items": [
      ["课程名称", "Python程序设计基础"],
      ["授课班级", "示例班级"],
      ["任课教师", "示例教师"],
      ["授课周数", "第 1 周 ～ 第 4 周"]
    ],
    "school": "示例职业学校",
    "date": "示例日期"
  },
  "schedule": {
    "title": "教学进度表",
    "stats": {
      "weeks_count": "20",
      "week_hours": "4/班",
      "total_hours": "80",
      "lecture_hours": "36",
      "lab_hours": "38",
      "exam_hours": "6"
    },
    "rows": [
      ["1", "理论教学内容简述", "2", "实操/实验教学内容简述", "2"],
      ["2", "理论教学内容简述", "2", "实操/实验教学内容简述", "2"]
    ],
    "notes": "备注说明事项"
  },
  "lessons": [
    {
      "lesson_num": "教案 1",
      "title": "课时课题名称（如：Linux系统认知与终端基础）",
      "hours": "2课时",
      "knowledge_obj": "1. 知识点1...\n2. 知识点2...",
      "ability_obj": "1. 技能点1...\n2. 技能点2...",
      "quality_obj": "1. 职业素养/工程规范...",
      "key_points": "本课核心重点（简明扼要）",
      "difficult_points": "学生实操认知卡点（具体）",
      "teaching_methods": "任务驱动法、理实一体化教学法、示范演示法",
      "learning_methods": "自主探究法、小组协作法、对比归纳法",
      "env_resources": "多媒体机房、CentOS/Ubuntu虚拟机、教学课件、实训工单",
      "stages": [
        {
          "stage_index": "（一）",
          "stage_name": "任务初探 / 导入新课",
          "design": "【活动1：情境创设】\n1. 教师活动：...\n2. 学生活动：...",
          "intent": "设计意图说明（简短精炼）"
        },
        {
          "stage_index": "（二）",
          "stage_name": "分析任务 / 讲授新知",
          "design": "【活动2：原理解析与示范】\n1. 教师活动：...\n2. 学生活动：...",
          "intent": "设计意图说明"
        },
        {
          "stage_index": "（三）",
          "stage_name": "任务执行 / 技能实操",
          "design": "【活动3：分组实训】\n1. 教师活动：...\n2. 学生活动：...",
          "intent": "强化动手实操能力"
        },
        {
          "stage_index": "（四）",
          "stage_name": "任务分享 / 成果展示",
          "design": "【活动4：成果互评】\n1. 教师活动：...\n2. 学生活动：...",
          "intent": "培养表达与反思习惯"
        },
        {
          "stage_index": "（五）",
          "stage_name": "任务评价 / 课堂小结",
          "design": "【活动5：总结提升】\n1. 教师活动：梳理知识逻辑与易错点...\n2. 学生活动：对照板书回顾...",
          "intent": "巩固知识主干"
        },
        {
          "stage_index": "（六）",
          "stage_name": "任务拓展 / 课后作业",
          "design": "【课后任务】\n1. 基础作业：...\n2. 拓展提升：...",
          "intent": "拓展工程视野"
        }
      ]
    }
  ]
}
```

---

## 2. 字段规范与去 AI 化要求

*   **`title`**：课题名称须体现明确的项目化任务或章节知识（如“FHS目录结构解析与路径寻址实操”）。
*   **`stages`**：
    *   `stage_index`：大写汉字括号序号，如 `（一）`、`（二）`。
    *   `stage_name`：环节名称，中职推荐使用“六环节”（任务初探、分析任务、任务执行、任务分享、任务评价、任务拓展），也可自适应为其他环节划分。
    *   `design`：必须使用 `【活动X：...】` 开头，内分 `1. 教师活动：` 与 `2. 学生活动：`。根据 `pedagogy_guidelines.md`，写实具体操作命令与卡点，严禁假大空。
    *   `intent`：说明该环节的教学法依据与育人目标，语言凝练（1-2句）。
