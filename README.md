# Teaching Plan - 智能教案设计与 Word 排版生成引擎

专业教案设计与 Word 高保真排版生成工具。支持通过自然语言对话（作为 **Agent Skill**）生成符合学校教务审查规范的教案，或通过本地命令行（作为 **CLI 工具**）从结构化 JSON 批量编译 `.docx` 教学设计方案。

[快速上手 (Skill)](#skill-usage) · [命令行使用 (CLI)](#cli-usage) · [核心特性](#features) · [目录结构](#structure) · [测试与评测](#testing) · [阶段文档](#docs)

---

## 🚀 快速上手 <a id="quick-start"></a>

### 方式一：作为 Agent Skill 使用（推荐） <a id="skill-usage"></a>

#### 1. 一键让 Agent 自动安装
将下面这句话直接发送给您的 AI Agent（如 Codex / Claude Code / Antigravity / 其他支持 Skills 的 Agent）：

```text
帮我安装这个skill：https://github.com/hkra0/teaching-plan
```

> Agent 会自动读取仓库并部署 `teaching-plan`。安装完成后，技能显示名为 **「教案设计与 Word 排版」**。

#### 2. 调用 Skill 生成教案
在对话中通过自然语言直接调用：

- **从课程要求新建教案**：
  ```text
  使用 $teaching-plan 根据我提供的课程信息设计教案并生成 Word 文档。
  ```
- **根据已有参考文档逆向排版**：
  ```text
  使用 $teaching-plan 参考附件中的《参考教案.docx》，为我生成《Python应用开发》前4周教学设计。
  ```

<details>
<summary><b>🛠️ 手动安装</b></summary>

<br>

##### 方法 A：运行一键安装脚本
克隆本仓库到本地后执行：
```bash
# 默认部署至通用 Agent 技能目录 (~/.agents/skills/teaching-plan，会自动探测本机已有的 Agent 目录)
python3 install_skill.py

# 安装到当前项目局部目录 (.agents/skills/teaching-plan)
python3 install_skill.py --local

# 或指定自定义安装路径
python3 install_skill.py --dest ~/.claude/skills
```

##### 方法 B：手动复制
将本仓库内的 `skills/teaching-plan` 文件夹完整复制到您的 Agent Skills 目录：
- **通用 Agent 全局目录**：`~/.agents/skills/teaching-plan`
- **Claude Code 目录**：`~/.claude/skills/teaching-plan`
- **项目局部目录**：当前工作区根目录下的 `.agents/skills/teaching-plan`

##### 安装核心依赖
```bash
pip install -r skills/teaching-plan/requirements.txt
```

</details>

---

### 方式二：作为本地 CLI 工具使用 <a id="cli-usage"></a>

无需 AI Agent 介入，直接在本地通过命令行将 JSON 内容文件编译为排版规范的 Word (`.docx`) 文档。

#### 1. 安装依赖
```bash
pip install -r skills/teaching-plan/requirements.txt
```

#### 2. 编译教案
```bash
# 编译单个科目教案（以公开示范课示例为例）
python3 main.py "skills/teaching-plan/examples/Python程序设计教学设计（示例）.json"

# 批量编译某个目录下的所有教案并指定输出目录
python3 main.py content/ -d dist/

# 指定特定排版模板
python3 main.py -t default_vocational lesson.json

# 查看完整 CLI 参数与使用说明
python3 main.py --help
```

---

## ✨ 核心特性 <a id="features"></a>

- 🎯 **写实教学法（去 AI 味）**：教学目标强制绑定可观察行为动词；环节活动清晰划分“教师动作、学生操作、技术卡点、验收标准”；杜绝空泛套话与虚构课堂细节。
- 📐 **高保真排版渲染引擎**：声明式模板与样式体系，精确控制页边距、表格列宽、跨页断表、标题孤行控制与技术代码/公式等宽对齐。
- 🔍 **双模工作流**：支持使用内置标准职教/普教模板从零生成，或基于已有 `.docx` 参考样例文档逆向提取页面与表格几何结构进行近似复现。
- 🛡️ **严密的质量保障体系**：内置 9 维教学质量量表（D1–D9）、8 大典型教学场景评测集（TC-01 ~ TC-08）与 76 项自动化几何与渲染测试。

---

## 📁 目录结构 <a id="structure"></a>

```text
teaching-plan/
├── skills/                     # 独立可发布的 Agent Skill 包
│   └── teaching-plan/
│       ├── SKILL.md            # Skill 规范定义与教学法 prompt
│       ├── scripts/            # 文档编译引擎、内容校验与排版工具
│       ├── templates/          # 声明式排版布局模板 (default_vocational.json 等)
│       ├── styles/             # 样式与字体配置 (default_styles.json)
│       ├── references/         # 教学法、Schema规范、运行时与量表参考
│       ├── examples/           # 公开示例 (已脱敏)
│       ├── evals/              # 教学质量评测用例、候选输出与双评审记录
│       ├── tests/              # 自动化测试套件 (76 项质量与回归测试)
│       ├── LICENSE             # MIT 开源许可证
│       └── requirements.txt    # 核心渲染依赖 (python-docx)
├── docs/                       # 设计阶段、评测基线与验收归档文档
│   ├── README.md               # 文档索引
│   ├── P0验收与下一步.md        # P0 渲染流水线与几何回归
│   ├── P1-2评测基线.md         # P1-2 教学质量 9 维量表与评测用例
│   ├── P1-2A发布清理.md        # P1-2A 脱敏清理与依赖收敛
│   └── P1-2B评测收口.md        # P1-2B 双独立评审收口记录
├── main.py                     # 本地 CLI 工具入口 (支持单文件与批量生成)
├── install_skill.py            # 一键将 Skill 安装到全局技能目录的脚本
└── .gitignore                  # Git 忽略配置 (过滤二进制产物、缓存与本地私有数据)
```

---

## 🧪 测试与质量评测 <a id="testing"></a>

### 运行自动化测试套件
```bash
python3 -m unittest discover -s skills/teaching-plan/tests
```
*(包含 76 项测试，覆盖排版布局、分页、跨页对齐、代码排版及内容 Schema 校验)*

### 运行教学质量评分校验
```bash
python3 skills/teaching-plan/scripts/evaluate_quality.py skills/teaching-plan/evals/reviews/p1-2b-round-2.json
```

---

## 📖 阶段文档 <a id="docs"></a>

项目各阶段的演进设计、技术规范与评测报告详见 [docs/ 目录](docs/README.md)：
- [P0 / P1-1 渲染流水线与几何回归](docs/P0验收与下一步.md)
- [P1-2 教学质量评测基线 (9 维量表与用例)](docs/P1-2评测基线.md)
- [P1-2A 发布清理与脱敏规范](docs/P1-2A发布清理.md)
- [P1-2B 独立双评审收口记录](docs/P1-2B评测收口.md)

---

## 📄 许可证

本项目基于 [MIT License](skills/teaching-plan/LICENSE) 开源。
