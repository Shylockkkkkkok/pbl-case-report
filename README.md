# pbl-case-report

skill：给一个 PBL 临床案例与需要汇报的问题，自动完成调研并产出 **1 份 Word 报告 + 1 份汇报 PPT**。

## 它做什么

1. 读案例（PDF / 文本），提取临床要素，并抠出案例里的查体照片、超声等原始图片复用
2. 从案例线索中定位研究问题
3. 调研（NCBI E-utilities，逐条核对 PMID；教材用于基础机制）
4. 生成 Word 报告（按「回答问题」组织，含「案例联系」一节）
5. 生成汇报 PPT（12–13 页 / 5 分钟，每页带参考文献脚注，末页参考文献 + 谢谢大家）
6. 版面自检（不启动 Office）

## 安装

标准的 `SKILL.md` 格式，**同一个文件夹在 Claude Code、Codex、OpenCode、Cursor、Gemini CLI 等 79 种 agent 中通用**。

### 推荐：skills CLI（npm 包 `skills`）

```bash
# 交互式：自己勾选装到哪些 agent（推荐）
npx skills add Shylockkkkkkok/pbl-case-report -g

# 非交互：装到所有检测到的 agent
# （少数仅支持项目级安装的 agent 会提示跳过，属正常，不影响其他 agent）
npx skills add Shylockkkkkkok/pbl-case-report -g --all

# 只装指定 agent
npx skills add Shylockkkkkkok/pbl-case-report -g -a claude-code -a codex -a opencode

# 只看仓库里有哪些 skill，不安装
npx skills add Shylockkkkkkok/pbl-case-report -l

# 在 skills.sh 上搜索
npx skills find pbl-case-report
```

需要 Node.js 22.20 或更高版本。

各 agent 的安装目录（各自独立）：

| Agent | 全局路径 |
|---|---|
| Claude Code | `~/.claude/skills/` |
| Codex / Cline / Gemini CLI / Cursor | `~/.agents/skills/`、`~/.gemini/skills/`、`~/.cursor/skills/` 等 |
| OpenCode | `~/.config/opencode/skills/` |

完整列表见 [docs/agents.md](https://github.com/vercel-labs/skills/blob/main/docs/agents.md)。

### 手动安装（只用 Claude Code）

```bash
git clone https://github.com/Shylockkkkkkok/pbl-case-report ~/.claude/skills/pbl-case-report
```

ZIP 解压、`cp -r`、软链接都可以。

## 结构

```
SKILL.md                  工作流、硬性排版要求、验收清单
reference/
  word_helpers.py         python-docx 排版库（含 Word 主题字体覆盖的修正）
  ppt_helpers.js          pptxgenjs 版式库（含防负尺寸、图注防重叠）
  figures.py              matplotlib 中文图表 + 机制通路图 + 脉冲线
  qa.py                   版面几何自检 + 纯 Python 预览渲染
```

## 依赖

- Python 3.10+：`python-docx` `python-pptx` `matplotlib` `Pillow` `pymupdf`
- Node：`pptxgenjs` `react-icons` `react` `react-dom` `sharp`
- 字体：思源黑体 / 微软雅黑 / Times New Roman

## 说明

仅为辅助PBL学习，具体内容请自行审核！


## 许可

[MIT License](LICENSE) —— 可以自由使用、修改、分发，包括商业用途，保留版权声明即可。

## 模板库

`templates/xjtu/` 内置 7 套西安交通大学 PPT 模板（已去重），默认使用 `飞上云霄.pptx`。
详见 [templates/xjtu/README.md](templates/xjtu/README.md)。

> 全部模板合计约 62 MB，克隆仓库时会一并下载。

## 内置第三方 skill

`skills/editable-research-ppt/` 原样收录自 [Daylily-Huang/ScholarFlow](https://github.com/Daylily-Huang/ScholarFlow)（v1.7.1），用于生成可编辑的科研 PPTX。

- 代码与模板：**MIT**，版权归 `editable-research-ppt contributors`，见其 `LICENSE`
- 字体：SIL OFL 1.1；图标与素材：ISC / CC0 / CC BY 等，逐项记录在 `THIRD_PARTY_NOTICES.md` 与 `assets/media-library/licenses/`

本仓库只做收录，未修改其内容。使用时以其自带 `SKILL.md` 为准。
