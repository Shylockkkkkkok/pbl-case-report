# -*- coding: utf-8 -*-
"""
matplotlib 中文图表库 —— 供 PPT / Word 内嵌

要点：
  · 中文必须显式指定可用字体，否则方块乱码
  · 输出透明背景 PNG，直接贴到浅色幻灯片上
  · 每张图保存后打印真实像素与宽高比 —— 填进 ppt_helpers.js 的 AR 表，
    这样 chartCard 才能把图注放在图片底边之下而不压图
  · ⚠️ 宽图放进竖向卡片会被高度限制而缩小，图内字会变小。
    宁可把 figsize 做得更宽（如 12×5.4），也不要事后靠调 PPT 字号补救。

用法：
    from figures import *
    setup()
    effect_chart([("指标 A", 1.82, 0.71, 2.94, True, "P < 0.001"),
                  ("指标 B", 0.49, 0.17, 0.81, True, "P = 0.027"),
                  ("指标 C", 0.28, -0.36, 0.93, False, "P = 0.264")],
                 "某结局的合并效应量", "assets/effect.png")
"""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
import numpy as np
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch

# ---------------- 配色（与 ppt_helpers.js 的 C 保持一致） ----------------
NIGHT = "#0E2233"
NIGHT2 = "#17384F"
TEAL = "#1C7293"
AMBER = "#E8871E"
RED = "#C0392B"
GREEN = "#0D9488"
PURPLE = "#7A4E9B"
GREY = "#64748B"
LIGHT = "#F2F6F9"


def setup():
    """全局中文字体。Windows 下这几款一定存在。"""
    plt.rcParams["font.sans-serif"] = ["Microsoft YaHei", "SimHei", "Source Han Sans CN"]
    plt.rcParams["axes.unicode_minus"] = False
    plt.rcParams["figure.dpi"] = 200


def _out(fig, path, transparent=True):
    fig.savefig(path, dpi=220, transparent=transparent, bbox_inches="tight")
    plt.close(fig)
    from PIL import Image
    im = Image.open(path)
    print("%-34s %4d x %4d   AR=%.3f" % (path, im.width, im.height, im.width / im.height))


def _clean(ax, keep=("bottom",)):
    for s in ("top", "right", "left", "bottom"):
        ax.spines[s].set_visible(s in keep)
    for s in keep:
        ax.spines[s].set_color("#D3DBE2")


# ============================================================
# 横向条形图 + 95% CI（Meta 分析效应量）
# ============================================================
def effect_chart(items, title_text, path, xlabel="Hedges' g",
                 figsize=(8.6, 4.3), transparent=True):
    """
    items: [(标签, 点估计, 下限, 上限, 是否显著, 右侧说明), ...]
           顺序即自上而下的展示顺序（传进来的第一项画在最上面）
    """
    fig, ax = plt.subplots(figsize=figsize)
    fig.patch.set_alpha(0 if transparent else 1)

    n = len(items)
    y = np.arange(n)[::-1]
    colors = [AMBER, TEAL, "#A9B6C2", PURPLE, GREEN]
    label_x = max(max(i[3] for i in items) * 1.06, 1.0)

    for k, (lab, g, lo, hi, sig, note) in enumerate(items):
        col = colors[k % len(colors)]
        ax.barh(y[k], g, height=0.46, color=col, alpha=0.92, zorder=3, edgecolor="none")
        if sig:
            ax.plot([lo, hi], [y[k], y[k]], color=col, lw=2.4, zorder=4, solid_capstyle="round")
            for x0 in (lo, hi):
                ax.plot([x0, x0], [y[k] - 0.13, y[k] + 0.13], color=col, lw=2.4, zorder=4)
        ax.text(label_x, y[k], note, va="center", fontsize=10.2,
                color=NIGHT if sig else "#8A97A3",
                fontweight="bold" if sig else "normal")

    ax.axvline(0, color="#8A97A3", lw=1.4, zorder=2)
    ax.set_yticks(y)
    ax.set_yticklabels([i[0] for i in items], fontsize=15, color=NIGHT, fontweight="bold")
    ax.set_xlim(min(0, min(i[2] for i in items)) - 1.2, label_x * 1.9)
    ax.set_xlabel(xlabel, fontsize=11.5, color=GREY, labelpad=8)
    ax.tick_params(axis="x", labelsize=10.5, colors=GREY)
    ax.tick_params(axis="y", length=0)
    _clean(ax)
    ax.set_title(title_text, fontsize=13.5, color=NIGHT, fontweight="bold", pad=14, loc="left")
    plt.tight_layout()
    _out(fig, path, transparent)


# ============================================================
# 柱状图 + 误差棒（风险比 / 组间比较）
# ============================================================
def bar_chart(labels, values, lo, hi, title_text, path, ylabel="",
              figsize=(7.2, 4.0), colors=None, transparent=True):
    fig, ax = plt.subplots(figsize=figsize)
    fig.patch.set_alpha(0 if transparent else 1)
    x = np.arange(len(values))
    cols = colors or ([TEAL] * len(values))
    ax.bar(x, values, width=0.56, color=cols, zorder=3, edgecolor="none")
    for i in range(len(values)):
        if lo[i] is not None and hi[i] is not None and (lo[i] != hi[i]):
            ax.plot([x[i], x[i]], [lo[i], hi[i]], color=NIGHT, lw=1.8, zorder=5)
            for yy in (lo[i], hi[i]):
                ax.plot([x[i] - 0.07, x[i] + 0.07], [yy, yy], color=NIGHT, lw=1.8, zorder=5)
        ax.text(x[i], hi[i] + 0.035 if hi[i] else values[i] + 0.035,
                "%.2f" % values[i], ha="center", fontsize=11.5, color=NIGHT, fontweight="bold")
    ax.axhline(1.0, color="#8A97A3", lw=1.2, ls=":", zorder=2)
    ax.set_xticks(x)
    ax.set_xticklabels(labels, fontsize=10.6, color=NIGHT)
    ax.set_ylabel(ylabel, fontsize=11, color=GREY, labelpad=8)
    ax.tick_params(axis="y", labelsize=10.5, colors=GREY)
    ax.tick_params(axis="x", length=0)
    _clean(ax, keep=("left", "bottom"))
    ax.set_title(title_text, fontsize=13.5, color=NIGHT, fontweight="bold", pad=14, loc="left")
    plt.tight_layout()
    _out(fig, path, transparent)


# ============================================================
# 昼夜节律曲线（示意图）
# ============================================================
def rhythm_chart(path, title_text, x0=8, x1=32, peak=27, sleep=(23, 31),
                 figsize=(8.2, 4.1), transparent=True):
    """x 轴取 8:00 → 次日 8:00，使睡眠时段连续居中。"""
    fig, ax = plt.subplots(figsize=figsize)
    fig.patch.set_alpha(0 if transparent else 1)
    t = np.linspace(x0, x1, 800)
    normal = 100 + 62 * np.exp(-((t - peak) ** 2) / 8) - 26 * np.exp(-((t - 17) ** 2) / 23)
    flat = 100 + 21 * np.exp(-((t - (peak + 3)) ** 2) / 46) - 11 * np.exp(-((t - 18) ** 2) / 26)

    ax.axvspan(sleep[0], sleep[1], color="#DCE6EE", alpha=0.9, zorder=0)
    ax.text((sleep[0] + sleep[1]) / 2, 62, "习惯性睡眠时段", ha="center", fontsize=9.8, color="#7C8B99")
    ax.plot(t, normal, color=TEAL, lw=3.2, zorder=4)
    ax.plot(t, flat, color=RED, lw=3.2, ls=(0, (5, 2.6)), zorder=4)

    ax.set_xlim(x0, x1); ax.set_ylim(55, 192)
    ticks = list(range(x0, x1 + 1, 4))
    ax.set_xticks(ticks)
    ax.set_xticklabels(["%d:00" % (h % 24) for h in ticks], fontsize=10.5, color=GREY)
    ax.set_yticks([])
    _clean(ax)
    ax.set_title(title_text, fontsize=13.5, color=NIGHT, fontweight="bold", pad=14, loc="left")
    plt.tight_layout()
    _out(fig, path, transparent)


# ============================================================
# 机制通路图（左：诱因 → 中：多条通路 → 右：结局 → 底部反馈环）
# ============================================================
def flow_diagram(path, left_title, left_sub, mech, right_title, right_sub,
                 feedback="反馈放大 · 恶性循环", footer="", figsize=(12.0, 5.4),
                 transparent=False):
    """
    mech: [(编号标题, 副标题, 颜色), ...]  建议 4~5 条
    """
    FW, FH = figsize
    fig, ax = plt.subplots(figsize=figsize, dpi=200)
    ax.set_xlim(0, FW); ax.set_ylim(0, FH); ax.axis("off")

    def box(x, y, w, h, fc, ec, lw=1.6, r=0.12, z=2):
        ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle=f"round,pad=0,rounding_size={r}",
                                    linewidth=lw, edgecolor=ec, facecolor=fc, zorder=z))

    def arrow(x1, y1, x2, y2, color, lw=1.9, rad=0.0, z=3, ms=13):
        ax.add_patch(FancyArrowPatch((x1, y1), (x2, y2), arrowstyle="-|>", mutation_scale=ms,
                                     linewidth=lw, color=color, zorder=z,
                                     connectionstyle=f"arc3,rad={rad}", shrinkA=1, shrinkB=1))

    LW_ = 2.70
    box(0.25, FH / 2 - 0.60, LW_, 1.20, NIGHT, NIGHT)
    ax.text(0.25 + LW_ / 2, FH / 2 + 0.25, left_title, ha="center", va="center",
            fontsize=16, color="white", fontweight="bold", zorder=4)
    ax.text(0.25 + LW_ / 2, FH / 2 - 0.23, left_sub, ha="center", va="center",
            fontsize=10.5, color="#BFD9E6", zorder=4)

    BX, BW = 3.30, 5.20
    n = len(mech)
    GAP = 0.10
    BH = min(0.78, (FH - 0.95 - GAP * (n - 1)) / n)
    top = FH - 0.15
    for i, (t, s, col) in enumerate(mech):
        y = top - (i + 1) * BH - i * GAP
        box(BX, y, BW, BH, "white", col, lw=1.7)
        box(BX, y, 0.075, BH, col, col, lw=0, r=0.0, z=3)
        ax.text(BX + 0.26, y + BH * 0.64, t, ha="left", va="center",
                fontsize=13, color=NIGHT, fontweight="bold", zorder=4)
        ax.text(BX + 0.26, y + BH * 0.25, s, ha="left", va="center",
                fontsize=10.2, color=GREY, zorder=4)
        arrow(0.25 + LW_ + 0.02, FH / 2, BX - 0.02, y + BH / 2, col, lw=1.6, rad=0.09, ms=11)

    RX = BX + BW + 0.45
    RW = FW - RX - 0.25
    box(RX, FH / 2 - 0.60, RW, 1.20, RED, RED)
    ax.text(RX + RW / 2, FH / 2 + 0.25, right_title, ha="center", va="center",
            fontsize=16, color="white", fontweight="bold", zorder=4)
    ax.text(RX + RW / 2, FH / 2 - 0.23, right_sub, ha="center", va="center",
            fontsize=10.5, color="#F6DCD8", zorder=4)

    for i, (_, _, col) in enumerate(mech):
        y = top - (i + 1) * BH - i * GAP
        arrow(BX + BW + 0.02, y + BH / 2, RX - 0.02, FH / 2, col, lw=1.6, rad=-0.09, ms=11)

    # 底部反馈环
    FB = 0.52
    ax.plot([RX + RW / 2, RX + RW / 2], [FH / 2 - 0.62, FB], color=RED, lw=2.1, zorder=1, solid_capstyle="round")
    ax.plot([RX + RW / 2, 0.25 + LW_ / 2], [FB, FB], color=RED, lw=2.1, zorder=1, solid_capstyle="round")
    arrow(0.25 + LW_ / 2, FB, 0.25 + LW_ / 2, FH / 2 - 0.62, RED, lw=2.1, ms=14, z=1)
    ax.text(FW / 2, FB, feedback, ha="center", va="center", fontsize=11.5, color=RED,
            fontweight="bold", style="italic", zorder=5,
            bbox=dict(boxstyle="round,pad=0.30", fc="white", ec=RED, lw=1.0))
    if footer:
        ax.text(FW / 2, 0.14, footer, ha="center", va="center", fontsize=11.5,
                color=RED, fontweight="bold", zorder=4)

    plt.tight_layout(pad=0.15)
    _out(fig, path, transparent)


# ============================================================
# 心电脉冲线（封面 / 结尾页视觉母题）
# ============================================================
def pulse_line(path, beats=9, color=AMBER, figsize=(13.4, 1.5)):
    fig, ax = plt.subplots(figsize=figsize, dpi=220)
    fig.patch.set_alpha(0)
    ax.set_facecolor("none")

    def beat(t0, amp):
        xs = np.array([0, 0.16, 0.26, 0.33, 0.40, 0.48, 0.58, 0.72, 1.0])
        ys = np.array([0, 0, 0.10, -0.16, 1.0, -0.30, 0, 0, 0]) * amp
        return t0 + xs, ys

    X, Y = [], []
    for k in range(beats):
        x, y = beat(k * 1.0, 1.0 if k % 2 == 0 else 0.92)
        X.append(x); Y.append(y)
    ax.plot(np.concatenate(X), np.concatenate(Y), color=color, lw=2.6,
            solid_capstyle="round", solid_joinstyle="round")
    ax.set_xlim(-0.1, beats + 0.1); ax.set_ylim(-0.55, 1.25); ax.axis("off")
    plt.tight_layout(pad=0)
    fig.savefig(path, dpi=220, transparent=True, bbox_inches="tight", pad_inches=0)
    plt.close(fig)
    from PIL import Image
    im = Image.open(path)
    print("%-34s %4d x %4d   AR=%.3f" % (path, im.width, im.height, im.width / im.height))


if __name__ == "__main__":
    setup()
    import os
    os.makedirs("assets", exist_ok=True)
    # ⚠️ 下面这段演示数据来自首个 PBL 案例（内分泌/睡眠），只用于验证函数能跑；
    #    换成你的题目时整段替换，不要沿用这里的术语和数值。
    effect_chart(
        [("fT3", 3.82, 1.71, 5.94, True, "g = 3.82  (1.71–5.94)  P < 0.001"),
         ("TSH", 1.49, 0.17, 2.81, True, "g = 1.49  (0.17–2.81)  P = 0.027"),
         ("fT4", 0.48, -0.36, 1.33, False, "g = 0.48  P = 0.264（无统计学意义）")],
        "睡眠剥夺后甲状腺指标的合并效应量", "assets/effect.png")
    flow_diagram(
        "assets/flow.png", "熬夜 / 睡眠剥夺", "睡眠不足 · 节律紊乱 · 夜间光照",
        [("①  生物钟–HPT 轴紊乱", "SCN 与甲状腺外周钟失同步；TSH 夜间峰值消失", TEAL),
         ("②  交感神经系统过度激活", "去甲肾上腺素↑、MSNA↑；协同放大 β 受体效应", "#2E8BA8"),
         ("③  HPA 轴与心理应激", "皮质醇↑；负性生活事件是发病与复发的诱因", AMBER),
         ("④  免疫紊乱与自身免疫", "促炎因子↑、Th17/Treg 失衡；产生 TRAb", RED),
         ("⑤  氧化应激与腺体损伤", "ROS↑、抗氧化能力↓；自噬/凋亡失衡", PURPLE)],
        "甲亢发生与加重", "高代谢 · 交感兴奋 · 自身免疫活化",
        footer="甲亢 → 心悸、焦虑、盗汗、失眠 → 睡眠进一步剥夺 → 甲状腺轴持续紊乱 → 症状加重")
    pulse_line("assets/pulse.png")
