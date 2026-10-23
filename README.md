# pbl-case-report

一个 Claude Code skill：给一个 PBL 临床案例，自动完成调研并产出 **1 份 Word 报告 + 1 份汇报 PPT**。

## 它做什么

1. 读案例（PDF / 文本），提取临床要素，并抠出案例里的查体照片、超声等原始图片复用
2. 从案例线索中定位研究问题
3. 调研（NCBI E-utilities，逐条核对 PMID；教材用于基础机制）
4. 生成 Word 报告（按「回答问题」组织，含「案例联系」一节）
5. 生成汇报 PPT（12–13 页 / 5 分钟，每页带参考文献脚注，末页参考文献 + 谢谢大家）
6. 版面自检（不启动 Office）

## 安装

```bash
git clone <this-repo> ~/.claude/skills/pbl-case-report
```

然后重启 Claude Code，`/skills` 里就能看到 `pbl-case-report`。

## 结构

```
SKILL.md                  工作流、硬性排版要求、踩过的坑、验收清单
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

`reference/` 里的四个文件都是可直接运行的代码，不是伪代码。`word_helpers.py`、`figures.py` 带 `__main__` 示例，可以直接跑一遍看效果。

## 注意

不要用 PowerPoint COM 导出图片——脚本结尾的 `Quit()` 会关掉整个 PowerPoint 进程，把正在编辑的演示文稿一起关掉。用 `qa.py`。
