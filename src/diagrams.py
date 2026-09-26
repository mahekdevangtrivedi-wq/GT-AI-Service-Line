# -*- coding: utf-8 -*-
"""Matplotlib diagrams shared by the framework document and strategy deck."""
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.patches import FancyBboxPatch, FancyArrowPatch
import numpy as np

GT = "#4F2D7F"; GT2 = "#7B5BA6"; LAV = "#EDE7F6"; TEAL = "#00A7B5"; GREY = "#5B5B5B"
RAG = {"Low": "#A9D18E", "Medium": "#FFD966", "High": "#F4B183", "Critical": "#E06666"}
plt.rcParams["font.family"] = "DejaVu Sans"


def process_diagram(path, n_risks=52):
    fig, ax = plt.subplots(figsize=(11, 4.2), dpi=200)
    ax.set_xlim(0, 110); ax.set_ylim(0, 47); ax.axis("off")
    stages = [("0", "Intake &\nContext", "Use-case profile,\nstakeholders, role,\nregulation"),
              ("1", "Screen &\nTier", "Prohibited check,\n10-factor tiering,\nLow to Critical"),
              ("2", "Identify", f"{n_risks}-risk taxonomy,\nthreat modelling,\nimpact assessment"),
              ("3", "Analyse", "Likelihood x impact\n(6 dimensions),\ninherent rating"),
              ("4", "Evaluate", "Control effectiveness,\nresidual risk vs\nappetite"),
              ("5", "Treat", "Avoid / mitigate /\ntransfer / accept;\nRCM controls"),
              ("6", "Monitor &\nReview", "KRIs, triggers,\nre-assessment,\nreporting")]
    w = 13.2; gap = 2.1; x0 = 1.5
    for i, (n, t, d) in enumerate(stages):
        x = x0 + i * (w + gap)
        col = GT if i % 2 == 0 else GT2
        ax.add_patch(FancyBboxPatch((x, 22), w, 15, boxstyle="round,pad=0.3,rounding_size=1.5", fc=col, ec="none"))
        ax.text(x + w / 2, 33.5, f"STAGE {n}", color="white", ha="center", va="center", fontsize=8, fontweight="bold", alpha=0.85)
        ax.text(x + w / 2, 27.5, t, color="white", ha="center", va="center", fontsize=9.5, fontweight="bold")
        ax.add_patch(FancyBboxPatch((x, 3), w, 16, boxstyle="round,pad=0.3,rounding_size=1.2", fc=LAV, ec=col, lw=1))
        ax.text(x + w / 2, 11, d, color="#222222", ha="center", va="center", fontsize=6.8)
        if i < len(stages) - 1:
            ax.annotate("", xy=(x + w + gap - 0.2, 29.5), xytext=(x + w + 0.4, 29.5),
                        arrowprops=dict(arrowstyle="-|>", color=GREY, lw=1.5))
    ax.annotate("", xy=(x0 + 3, 38.8), xytext=(x0 + 6 * (w + gap) + w - 3, 38.8),
                arrowprops=dict(arrowstyle="-|>", color=TEAL, lw=1.4, connectionstyle="arc3,rad=0.12"))
    ax.text(55, 44.5, "Continuous: communication & consultation  |  recording & reporting  |  re-assessment on material change", ha="center", fontsize=7.5, color=TEAL, style="italic")
    fig.savefig(path, bbox_inches="tight", facecolor="white"); plt.close(fig)


def heatmap(path):
    fig, ax = plt.subplots(figsize=(6.2, 4.6), dpi=200)
    L = ["Rare", "Unlikely", "Possible", "Likely", "Almost\ncertain"]
    I = ["Insignificant", "Minor", "Moderate", "Major", "Severe"]
    for li in range(5):
        for ii in range(5):
            s = (li + 1) * (ii + 1)
            band = "Critical" if s >= 17 else "High" if s >= 10 else "Medium" if s >= 5 else "Low"
            ax.add_patch(plt.Rectangle((ii, li), 1, 1, fc=RAG[band], ec="white", lw=2))
            ax.text(ii + 0.5, li + 0.5, str(s), ha="center", va="center", fontsize=11, fontweight="bold", color="#222")
    ax.set_xlim(0, 5); ax.set_ylim(0, 5)
    ax.set_xticks(np.arange(5) + 0.5); ax.set_xticklabels([f"{i+1}\n{x}" for i, x in enumerate(I)], fontsize=8)
    ax.set_yticks(np.arange(5) + 0.5); ax.set_yticklabels([f"{x} {i+1}" for i, x in enumerate(L)], fontsize=8)
    ax.set_xlabel("IMPACT", fontsize=9, fontweight="bold", color=GT); ax.set_ylabel("LIKELIHOOD", fontsize=9, fontweight="bold", color=GT)
    for s in ax.spines.values(): s.set_visible(False)
    ax.tick_params(length=0)
    handles = [plt.Rectangle((0, 0), 1, 1, fc=RAG[k]) for k in RAG]
    ax.legend(handles, ["Low (1-4)", "Medium (5-9)", "High (10-16)", "Critical (17-25)"], loc="upper center", bbox_to_anchor=(0.5, -0.2), ncol=4, fontsize=7.5, frameon=False)
    fig.savefig(path, bbox_inches="tight", facecolor="white"); plt.close(fig)


def lines_of_defence(path):
    fig, ax = plt.subplots(figsize=(10, 3.6), dpi=200)
    ax.set_xlim(0, 100); ax.set_ylim(-4, 36); ax.axis("off")
    ax.add_patch(FancyBboxPatch((1, 29), 98, 6, boxstyle="round,pad=0.2,rounding_size=1", fc=GT, ec="none"))
    ax.text(50, 32, "Board / Board Risk Committee  -  sets AI risk appetite, oversees Critical risks", color="white", ha="center", va="center", fontsize=9.5, fontweight="bold")
    ax.add_patch(FancyBboxPatch((1, 21), 98, 6, boxstyle="round,pad=0.2,rounding_size=1", fc=GT2, ec="none"))
    ax.text(50, 24, "ExCo & AI Committee / CoE  -  approves High & Critical use cases; owns guardrails", color="white", ha="center", va="center", fontsize=9.5, fontweight="bold")
    cols = [("1st line", "Use-case owners, product & data\nscience teams, IT / engineering", "Own and manage AI risks; perform\nassessments; operate controls"),
            ("2nd line", "Risk, Compliance, Privacy (DPO),\nInfoSec, Model validation", "Set framework & standards; challenge;\nvalidate; monitor KRIs; report"),
            ("3rd line", "Internal Audit", "Independent assurance over AI\ngovernance and control effectiveness")]
    for i, (h, who, what) in enumerate(cols):
        x = 1 + i * 33
        ax.add_patch(FancyBboxPatch((x, 1), 31.5, 17.5, boxstyle="round,pad=0.2,rounding_size=1", fc=LAV, ec=GT, lw=1))
        ax.text(x + 15.75, 15.8, h.upper(), ha="center", fontsize=10, fontweight="bold", color=GT)
        ax.text(x + 15.75, 10.6, who, ha="center", va="center", fontsize=7.2, color="#222", fontweight="bold")
        ax.text(x + 15.75, 4.6, what, ha="center", va="center", fontsize=7.2, color="#333")
    ax.text(99, -2.8, "External assurance: ISO/IEC 42001 certification bodies, regulators (CBB, PDPA), independent auditors", ha="right", fontsize=7, color=GREY, style="italic")
    fig.savefig(path, bbox_inches="tight", facecolor="white"); plt.close(fig)


def radar(path, labels, values, bench=None, title=None):
    n = len(labels)
    ang = np.linspace(0, 2 * np.pi, n, endpoint=False).tolist(); ang += ang[:1]
    v = list(values) + [values[0]]
    fig = plt.figure(figsize=(5.4, 5.0), dpi=200)
    ax = plt.subplot(111, polar=True)
    ax.set_theta_offset(np.pi / 2); ax.set_theta_direction(-1)
    ax.set_ylim(0, 100); ax.set_yticks([25, 50, 70, 90]); ax.set_yticklabels(["25", "50", "70", "90"], fontsize=7, color=GREY)
    ax.set_xticks(ang[:-1]); ax.set_xticklabels(labels, fontsize=8.5, color="#222")
    if bench:
        b = list(bench) + [bench[0]]
        ax.plot(ang, b, color=TEAL, lw=1.4, ls="--", label="Target (end-2027)")
    ax.plot(ang, v, color=GT, lw=2, label="Current (Sep 2026)"); ax.fill(ang, v, color=GT, alpha=0.18)
    for a, val in zip(ang[:-1], values):
        ax.text(a, val + 7, f"{val:.0f}%", ha="center", va="center", fontsize=8, fontweight="bold", color=GT)
    ax.legend(loc="lower center", bbox_to_anchor=(0.5, -0.16), ncol=2, fontsize=8, frameon=False)
    if title: ax.set_title(title, fontsize=10, color=GT, fontweight="bold", pad=18)
    fig.savefig(path, bbox_inches="tight", facecolor="white", transparent=False); plt.close(fig)


def pillar_bars(path, labels, values):
    fig, ax = plt.subplots(figsize=(7.2, 3.9), dpi=200)
    y = np.arange(len(labels))[::-1]
    for band, lo, hi, col in [("Partial", 0, 50, "#F8CBAD"), ("Informed", 50, 70, "#FFE699"), ("Repeatable", 70, 90, "#C5E0B4"), ("Adaptive", 90, 100, "#9BC2E6")]:
        ax.axvspan(lo, hi, color=col, alpha=0.45, lw=0)
        ax.text((lo + hi) / 2, len(labels) - 0.35, band, ha="center", fontsize=7.5, color="#444")
    bars = ax.barh(y, values, color=GT, height=0.55)
    for yy, v in zip(y, values):
        ax.text(v + 1.2, yy, f"{v:.0f}%", va="center", fontsize=9, fontweight="bold", color=GT)
    ax.set_yticks(y); ax.set_yticklabels(labels, fontsize=9)
    ax.set_xlim(0, 100); ax.set_ylim(-0.6, len(labels) - 0.1)
    ax.axvline(61.4, color=TEAL, ls="--", lw=1.2); ax.text(61.9, -0.55, "Overall 61.4%", color=TEAL, fontsize=8)
    for s in ["top", "right"]: ax.spines[s].set_visible(False)
    ax.tick_params(axis="x", labelsize=8)
    fig.savefig(path, bbox_inches="tight", facecolor="white"); plt.close(fig)


def gantt(path, tasks, months=18, start_label="Q4 2026"):
    """tasks: list of (workstream, name, start_month(0-based), duration, colour)"""
    fig, ax = plt.subplots(figsize=(12, 0.34 * len(tasks) + 1.6), dpi=200)
    qlabels = ["Q4-26", "Q1-27", "Q2-27", "Q3-27", "Q4-27", "Q1-28"]
    for q in range(0, months, 3):
        ax.axvspan(q, q + 3, color="#F4F1F9" if (q // 3) % 2 == 0 else "white", lw=0)
        ax.text(q + 1.5, -0.9, qlabels[q // 3], ha="center", fontsize=8.5, fontweight="bold", color=GT)
    for i, (ws, name, s, d, col) in enumerate(tasks):
        ax.barh(i, d, left=s, color=col, height=0.62, edgecolor="white")
        ax.text(-0.2, i, name, ha="right", va="center", fontsize=7.8, color="#222")
    ax.set_ylim(len(tasks) - 0.4, -1.4); ax.set_xlim(0, months)
    ax.set_yticks([]); ax.set_xticks([])
    for s in ax.spines.values(): s.set_visible(False)
    for m, lab in [(3, "Horizon 1 exit:\nfoundations"), (9, "Horizon 2 exit:\nscale"), (18, "Horizon 3:\ndifferentiate")]:
        ax.axvline(m, color=TEAL, ls=":", lw=1.2)
    fig.savefig(path, bbox_inches="tight", facecolor="white"); plt.close(fig)


def swot(path, s, w, o, t):
    fig, ax = plt.subplots(figsize=(11, 6.2), dpi=200)
    ax.set_xlim(0, 100); ax.set_ylim(0, 62); ax.axis("off")
    quads = [("STRENGTHS", s, 0.5, 31.5, GT), ("WEAKNESSES", w, 50.5, 31.5, GT2), ("OPPORTUNITIES", o, 0.5, 0.5, "#00838F"), ("THREATS", t, 50.5, 0.5, "#B23A48")]
    for title, items, x, y, col in quads:
        ax.add_patch(FancyBboxPatch((x, y), 49, 30, boxstyle="round,pad=0.2,rounding_size=1.2", fc="white", ec=col, lw=1.6))
        ax.add_patch(FancyBboxPatch((x, y + 25.2), 49, 4.8, boxstyle="round,pad=0.2,rounding_size=1.2", fc=col, ec=col))
        ax.text(x + 2, y + 27.6, title, color="white", fontsize=11, fontweight="bold", va="center")
        for k, it in enumerate(items):
            ax.text(x + 2, y + 22.5 - k * 4.1, "• " + it, fontsize=7.9, va="top", color="#222", wrap=True)
    fig.savefig(path, bbox_inches="tight", facecolor="white"); plt.close(fig)
