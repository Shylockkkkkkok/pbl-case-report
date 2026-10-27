# pbl-case-report

skill：给一个 PBL 临床案例与需要汇报的问题，自动完成调研并产出 **1 份 Word 报告 + 1 份汇报 PPT**。

## 它做什么

1. 读案例（PDF / 文本），提取临床要素，并抠出案例里的查体照片、超声等原始图片复用
2. 从案例线索中定位研究问题
3. 调研（NCBI E-utilities，逐条核对 PMID；教材用于基础机制）
4. 生成 Word 报告（按「回答问题」组织，含「案例联系」一节）
5. 生成汇报 PPT（引擎出稿 → 学校封面 → 逐页引用脚注，5 分钟节奏）
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
SKILL.md                       工作流、硬性排版要求、验收清单
LICENSE                        MIT
reference/
  word_helpers.py              python-docx 排版库（主题字体修正、参考文献重排与校验）
  ppt_helpers.js               pptxgenjs 版式库（防负尺寸、图注防重叠）
  figures.py                   matplotlib 中文图表 + 机制通路图 + 脉冲线
  qa.py                        版面几何自检 + 预览渲染（支持 srcRect 裁剪）
  apply_school_skin.py         套学校封面/致谢（--cover-only、--logo-scale）
  add_school_logo.py           只贴校徽（--scale、--only）
  add_ref_footers.py           逐页页脚引文（可按页覆盖位置与字号）
templates/xjtu/                3 套 16:9 交大模板 + 选型说明
.vendor/editable-research-ppt/ 内置 PPT 引擎（MIT，1169 个文件，SHA256 可校验）
```

## PPT 是怎么生成的

1. 读 `.vendor/editable-research-ppt/SKILL.md`，按「场景 → 配色 → 填充」确定方案
2. 写 `input.json`（逐页 `layout` L1–L23，`demo: false`），依次跑
   `preflight.py` → `generate.py` → `audit.py --strict`（必须 0 warning）
3. `add_ref_footers.py` 给引用页补页脚文献
4. `apply_school_skin.py --cover-only` 换学校封面；或 `add_school_logo.py` 只贴校徽
5. `qa.py` 复检版面（越界 / 溢出 / 重叠 / 引文编号）

## 依赖

- Python 3.10+：`python-docx` `python-pptx` `matplotlib` `Pillow` `pymupdf`
- Node：`pptxgenjs` `react-icons` `react` `react-dom` `sharp`
- 字体：思源黑体 / 微软雅黑 / Times New Roman

## 说明

仅为辅助PBL学习，具体内容请自行审核！


## 许可

[MIT License](LICENSE) —— 可以自由使用、修改、分发，包括商业用途，保留版权声明即可。

## 模板库

`templates/xjtu/` 内置 **3 套**西安交通大学 16:9 模板，**按主题选，不设默认**：

| 文件 | 适用场景 |
|---|---|
| `蓝色模板.pptx` | 临床 / 基础医学学术汇报（蓝白几何） |
| `西安交通大学汇报答辩通用模板.pptx` | 正式开题、毕业答辩（校红，版式最全） |
| `2022 西安交通大学创新港风格模板.pptx` | 工科、创新港相关 |

选型细则见 [templates/xjtu/README.md](templates/xjtu/README.md)。

## 内置 PPT 引擎

`.vendor/editable-research-ppt/` 原样收录自 [Daylily-Huang/ScholarFlow](https://github.com/Daylily-Huang/ScholarFlow)（v1.7.1），用于生成可编辑的科研 PPTX。

- 代码与模板：**MIT**，版权归 `editable-research-ppt contributors`，见其 `LICENSE`
- 字体：SIL OFL 1.1；图标与素材：ISC / CC0 / CC BY 等，逐项记录在 `THIRD_PARTY_NOTICES.md` 与 `assets/media-library/licenses/`
