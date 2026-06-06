import pandas as pd
import matplotlib.pyplot as plt
import japanize_matplotlib
import numpy as np

df = pd.read_csv("課題3 (2).csv")

COLORS = {
    "営業": "#4C72B0",
    "開発": "#DD8452",
    "管理": "#55A868",
    "人事": "#C44E52",
}
groups = df["所属"].unique()
color_list = [COLORS.get(g, "#8172B2") for g in groups]

# --- graph1: ドーナツ型円グラフ ---
counts = df["所属"].value_counts()
fig, ax = plt.subplots(figsize=(7, 7))
wedges, texts, autotexts = ax.pie(
    counts,
    labels=counts.index,
    autopct="%1.1f%%",
    startangle=90,
    colors=[COLORS.get(g, "#8172B2") for g in counts.index],
    wedgeprops=dict(width=0.5),
    pctdistance=0.75,
)
for t in autotexts:
    t.set_fontsize(11)
ax.set_title("所属別参加者数", fontsize=16, fontweight="bold", pad=20)
plt.tight_layout()
plt.savefig("graph1_pie.png", dpi=150, bbox_inches="tight")
plt.close()
print("graph1_pie.png 保存完了")

# --- graph2: 所属別スコア比較グループ棒グラフ（平均・最高・最低）---
stats = df.groupby("所属")["スコア"].agg(["mean", "max", "min"]).reset_index()
stats.columns = ["所属", "平均", "最高", "最低"]

x = np.arange(len(stats))
width = 0.25
fig, ax = plt.subplots(figsize=(9, 6))

bars_mean = ax.bar(x - width, stats["平均"], width, label="平均", color="#4C72B0")
bars_max  = ax.bar(x,          stats["最高"], width, label="最高", color="#DD8452")
bars_min  = ax.bar(x + width,  stats["最低"], width, label="最低", color="#55A868")

for bars in [bars_mean, bars_max, bars_min]:
    for bar in bars:
        ax.text(
            bar.get_x() + bar.get_width() / 2,
            bar.get_height() + 0.3,
            f"{bar.get_height():.1f}",
            ha="center", va="bottom", fontsize=8,
        )

ax.set_xticks(x)
ax.set_xticklabels(stats["所属"], fontsize=12)
ax.set_ylabel("スコア", fontsize=12)
ax.set_title("所属別スコア比較（平均・最高・最低）", fontsize=15, fontweight="bold")
ax.legend(fontsize=11)
ax.set_ylim(0, max(stats["最高"]) * 1.15)
ax.yaxis.grid(True, linestyle="--", alpha=0.7)
ax.set_axisbelow(True)
plt.tight_layout()
plt.savefig("graph2_bar.png", dpi=150, bbox_inches="tight")
plt.close()
print("graph2_bar.png 保存完了")

# --- graph3: 所属別色分けヒストグラム＋全体平均線 ---
fig, ax = plt.subplots(figsize=(9, 6))
bins = np.arange(65, 100, 5)

for group in df["所属"].unique():
    data = df[df["所属"] == group]["スコア"]
    ax.hist(
        data, bins=bins, alpha=0.6,
        label=group, color=COLORS.get(group, "#8172B2"),
        edgecolor="white",
    )

overall_mean = df["スコア"].mean()
ax.axvline(overall_mean, color="black", linestyle="--", linewidth=2,
           label=f"全体平均: {overall_mean:.1f}")

ax.set_xlabel("スコア", fontsize=12)
ax.set_ylabel("人数", fontsize=12)
ax.set_title("所属別スコア分布（ヒストグラム）", fontsize=15, fontweight="bold")
ax.legend(fontsize=11)
ax.yaxis.grid(True, linestyle="--", alpha=0.7)
ax.set_axisbelow(True)
plt.tight_layout()
plt.savefig("graph3_histogram.png", dpi=150, bbox_inches="tight")
plt.close()
print("graph3_histogram.png 保存完了")
