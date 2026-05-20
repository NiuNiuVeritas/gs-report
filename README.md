# GS Report · 国信证券周报转公众号 Skill

把国信证券的 Word 研报（`.docx`）一键转成微信公众号草稿（Markdown + HTML），自动套用「量化藏经阁」公众号的 `gzh` 模板：保留全部正文、自动抽取发布日期与分析师信息、生成图表占位、附带免责声明图片，并提供完整性校验。

这是一个 **Claude Code Skill**（带 `SKILL.md`），可以让 Claude 直接按工作流调用；脚本本身也能独立用 Python 跑。

---

## 一、它能做什么

输入：一份国信证券 Word 研报 `.docx`

输出：
- `xxx.md` — 公众号 Markdown（含内联 HTML 样式）
- `xxx.html` — 独立预览页（浏览器双击打开就能看效果）
- `assets/law.png` — 自动复制好的免责声明配图
- 校验报告 — 检查正文段落、图表标记、占位符、页脚元数据是否齐全

工作流细节见 [SKILL.md](SKILL.md) 和 [references/requirements.md](references/requirements.md)。

---

## 二、安装

本节提供两种安装方式：方式 A 由 Claude Code 自动完成，方式 B 为手动安装。任选其一。

前置要求：本机已安装 Git 与 Python 3.9+。Windows 用户可从 [Git for Windows](https://git-scm.com/download/win) 与 [python.org](https://www.python.org/downloads/) 安装（Python 安装时请勾选 "Add to PATH"）。

### 方式 A：由 Claude Code 自动安装

打开 Claude Code，向其发送以下指令：

```
请安装 gs-report 这个 skill：https://github.com/NiuNiuVeritas/gs-report.git
```

Claude 将自动完成：克隆仓库至用户级 skill 目录（`~/.claude/skills/gs-report/`）、读取本 README 中的安装步骤、安装 Python 依赖、并验证安装结果。安装完成后跳至 [三、使用](#三使用)。

> 仓库下方的 [给 Agent 的安装指引](#附给-agent-的安装指引) 提供了供 Claude 执行的标准化步骤，无需用户手动展开。

---

### 方式 B：手动安装

#### 1. 拉取代码

把这个仓库放到 Claude Code 的 skill 目录下。

**Windows（PowerShell）：**
```powershell
cd $env:USERPROFILE\.claude\skills
git clone https://github.com/NiuNiuVeritas/gs-report.git
```

**macOS / Linux：**
```bash
cd ~/.claude/skills
git clone https://github.com/NiuNiuVeritas/gs-report.git
```

> 不会用 git 也可以：在 GitHub 页面点 **Code → Download ZIP**，解压后把整个 `gs-report` 文件夹放到上面的 `skills` 目录里。注意解压后文件夹名必须叫 `gs-report`，不要带 `-main` 后缀。

> `~/.claude/skills/` 目录不存在的话，先手动建一下：Windows `mkdir $env:USERPROFILE\.claude\skills`，macOS/Linux `mkdir -p ~/.claude/skills`。

放好后路径应该长这样：
```
~/.claude/skills/gs-report/
  ├── SKILL.md
  ├── scripts/
  ├── assets/
  └── ...
```

#### 2. 安装 Python 依赖

需要 Python 3.9 及以上。脚本依赖两个库：

```bash
pip install python-docx lxml
```

> 用国内镜像更快：`pip install python-docx lxml -i https://pypi.tuna.tsinghua.edu.cn/simple`

#### 3. 确认 Claude Code 能识别

打开 Claude Code，输入：
```
/skills
```
列表里应该能看到 **gs-report**。看不到就检查路径是不是 `~/.claude/skills/gs-report/SKILL.md`。

---

## 三、使用

### 方式 A：让 Claude Code 直接帮你转（推荐）

最简单的用法 —— 把 docx 拖到对话框，或者直接说：

```
帮我用 gs-report 把这份周报转成公众号草稿：
D:\下载\国信证券-金融工程周报-XXX-20260520.docx
```

Claude 会自动：
1. 调用 `scripts/convert_gs_report.py` 生成 `.md` 和 `.html`
2. 调用 `scripts/verify_gs_report.py` 校验完整性
3. 报告输出路径和校验结果

### 方式 B：自己手动跑脚本

进到 skill 目录，跑：

```bash
python scripts/convert_gs_report.py --docx "你的报告.docx" --output-dir outputs
```

**常用参数：**

| 参数 | 说明 |
|---|---|
| `--docx` | 必填，输入的 Word 报告路径 |
| `--output-dir` | 输出目录（默认当前目录） |
| `--publication-date` | 手动指定发布日期，例如 `2026年04月28日`。一般不用填，脚本会自动从首页抽 |
| `--analyst` | 手动指定分析师，格式 `"张欣慰 S0980520060001"`，可以多次传 |
| `--slug` | 自定义输出文件名前缀 |

**示例（绝大多数情况这一行就够了）：**
```bash
python scripts/convert_gs_report.py --docx report.docx --output-dir outputs
```

**示例（自动抽取失败时再补参数）：**
```bash
python scripts/convert_gs_report.py --docx report.docx --output-dir outputs `
  --publication-date 2026年04月28日 `
  --analyst "张欣慰 S0980520060001" `
  --analyst "彭思宇 S0980521060003"
```

### 跑完一定要校验

```bash
python scripts/verify_gs_report.py --docx report.docx --markdown outputs/xxx.md
```

校验会检查：正文段落数 / 图表标记 / 占位符 / 页脚元数据 / 免责声明图片是否齐全。**有缺失就别发，回去看 Word 原文对照修。**

---

## 四、把生成结果发到公众号

1. 用浏览器打开 `outputs/xxx.html` 检查排版是不是符合预期
2. 打开公众号后台 → 新建图文
3. 把 `xxx.md` 内容粘贴进**支持 Markdown 的编辑器**（比如 [Md2All](http://md.aclickall.com/)、墨滴），渲染后复制到公众号编辑框
4. 图表占位行下面，**手动把 Word 里的图/表截图插入**（团队约定就是手工截图，脚本只生成占位）
5. 文末的 `law.png` 在 `outputs/assets/` 下，记得也插入到对应位置

---

## 五、常见问题

**Q：不熟悉命令行如何安装？**
使用 [安装方式 A](#方式-a由-claude-code-自动安装)，由 Claude Code 完成全部安装步骤。

**Q：`/skills` 里看不到 gs-report？**
确认路径是 `~/.claude/skills/gs-report/SKILL.md`（Windows 下是 `C:\Users\<你>\.claude\skills\gs-report\SKILL.md`）。文件夹名必须是 `gs-report`，不要 `gs-report-main`。

**Q：`ModuleNotFoundError: No module named 'docx'`？**
没装依赖。装的是 `python-docx`，不是 `docx`：
```bash
pip install python-docx lxml
```

**Q：发布日期或分析师抽不到？**
脚本会报警告。用 `--publication-date` 和 `--analyst` 手动补上重跑一次。

**Q：校验提示「正文段落缺失」？**
说明 Word 里有正文段落没被识别。先别发布，把 docx 和生成的 md 一起发给维护者，或者用 Claude Code 让 Claude 对照原文检查。

**Q：Word 文件锁定 / 临时文件 `~$xxx.docx`？**
那是 Word 打开文件时生成的临时锁文件，跟脚本无关。转之前把 Word 关掉就行。

---

## 六、目录结构

```
gs-report/
├── SKILL.md                 # Skill 工作流（Claude 读这个）
├── README.md                # 本文件
├── scripts/
│   ├── convert_gs_report.py # 转换脚本
│   └── verify_gs_report.py  # 校验脚本
├── assets/
│   ├── base-template.md     # 公众号模板样式
│   └── law.png              # 免责声明图
├── references/
│   └── requirements.md      # 转换规则细则
├── agents/
│   └── openai.yaml          # Agent 配置
└── outputs/                 # 生成结果默认放这
```

---

## 七、反馈

遇到问题或者有改进建议，提 [Issue](https://github.com/NiuNiuVeritas/gs-report/issues) 或者直接找维护者。

---

## 附：给 Agent 的安装指引

以下步骤供 Claude Code 在执行 [安装方式 A](#方式-a由-claude-code-自动安装) 时遵循，普通用户无需阅读。

1. **确定目标路径**：用户级 skill 目录为 `~/.claude/skills/gs-report/`（Windows: `%USERPROFILE%\.claude\skills\gs-report\`）。若父目录 `~/.claude/skills/` 不存在，先创建。
2. **获取代码**：若目标目录不存在，执行 `git clone https://github.com/NiuNiuVeritas/gs-report.git` 至该路径；若已存在且为 git 仓库，执行 `git pull` 更新；若已存在但非 git 仓库，向用户确认后再处理，不得直接覆盖。
3. **安装 Python 依赖**：执行 `pip install python-docx lxml`。若用户位于中国大陆网络环境，可追加 `-i https://pypi.tuna.tsinghua.edu.cn/simple`。
4. **验证安装**：确认 `<目标路径>/SKILL.md` 与 `<目标路径>/scripts/convert_gs_report.py` 存在，且 `python -c "import docx, lxml"` 无报错。
5. **汇报结果**：向用户报告安装路径、依赖安装状态、验证结果，并提示其在 Claude Code 内输入 `/skills` 确认 `gs-report` 已被识别。

如任一步骤失败，停止后续步骤并向用户说明失败原因与建议处置方式，不得静默重试或采取破坏性操作（如删除已有目录）。
