import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as ticker
import japanize_matplotlib
import numpy as np

df = pd.read_csv("課題2 (1).csv")
df["日付"] = pd.to_datetime(df["日付"])

COLORS = {
    "高橋花子": "#e05c5c",
    "鈴木花子": "#4C72B0",
    "田中一郎": "#55A868",
    "山田太郎": "#DD8452",
    "佐藤次郎": "#8172B2",
}

# --- graph4: 成長トレンド折れ線グラフ ---
daily = df.groupby(["名前", "日付"])["スコア"].mean().reset_index()
overall = df.groupby("日付")["スコア"].mean().reset_index()

fig, ax = plt.subplots(figsize=(10, 6))

for name, grp in daily.groupby("名前"):
    grp = grp.sort_values("日付")
    ax.plot(grp["日付"], grp["スコア"], marker="o", linewidth=2,
            label=name, color=COLORS.get(name))
    last = grp.iloc[-1]
    ax.annotate(f"{name}\n{last['スコア']:.1f}点",
                xy=(last["日付"], last["スコア"]),
                xytext=(8, 0), textcoords="offset points",
                fontsize=8, color=COLORS.get(name), va="center")

ax.plot(overall["日付"], overall["スコア"], color="black",
        linewidth=2.5, linestyle="--", marker="s", label="全体平均", zorder=5)

ax.set_title("日別スコア成長トレンド（受講者別）", fontsize=15, fontweight="bold")
ax.set_xlabel("日付", fontsize=12)
ax.set_ylabel("平均スコア", fontsize=12)
ax.set_ylim(60, 100)
ax.yaxis.set_major_locator(ticker.MultipleLocator(5))
ax.yaxis.grid(True, linestyle="--", alpha=0.6)
ax.set_axisbelow(True)
ax.legend(fontsize=10, loc="upper left")
fig.autofmt_xdate()
plt.tight_layout()
plt.savefig("graph4_trend.png", dpi=150, bbox_inches="tight")
plt.close()
print("graph4_trend.png 保存完了")

# --- graph5: 人物×科目 ヒートマップ ---
pivot = df.pivot_table(index="名前", columns="科目", values="スコア", aggfunc="mean")
# 平均でソート（高い順）
pivot = pivot.loc[pivot.mean(axis=1).sort_values(ascending=False).index]

fig, ax = plt.subplots(figsize=(9, 5))
im = ax.imshow(pivot.values, cmap="RdYlGn", aspect="auto", vmin=60, vmax=100)

ax.set_xticks(range(len(pivot.columns)))
ax.set_xticklabels(pivot.columns, fontsize=12)
ax.set_yticks(range(len(pivot.index)))
ax.set_yticklabels(pivot.index, fontsize=12)

for i in range(len(pivot.index)):
    for j in range(len(pivot.columns)):
        val = pivot.values[i, j]
        if np.isnan(val):
            text = "－"
            color = "gray"
        else:
            text = f"{val:.1f}"
            color = "black" if 72 < val < 93 else "white"
        ax.text(j, i, text, ha="center", va="center", fontsize=11,
                fontweight="bold", color=color)

plt.colorbar(im, ax=ax, label="スコア")
ax.set_title("受講者×科目 スコアヒートマップ", fontsize=15, fontweight="bold")
plt.tight_layout()
plt.savefig("graph5_heatmap.png", dpi=150, bbox_inches="tight")
plt.close()
print("graph5_heatmap.png 保存完了")

# --- graph6: 受講者別スコア分布（箱ひげ図）+ 成長幅バー ---
fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(13, 6))

names_sorted = df.groupby("名前")["スコア"].mean().sort_values(ascending=False).index.tolist()
data_by_name = [df[df["名前"] == n]["スコア"].values for n in names_sorted]
colors_list = [COLORS.get(n) for n in names_sorted]

bp = ax1.boxplot(data_by_name, patch_artist=True, vert=True,
                 medianprops=dict(color="black", linewidth=2))
for patch, color in zip(bp["boxes"], colors_list):
    patch.set_facecolor(color)
    patch.set_alpha(0.7)

ax1.set_xticklabels(names_sorted, fontsize=11)
ax1.set_ylabel("スコア", fontsize=12)
ax1.set_title("受講者別スコア分布", fontsize=14, fontweight="bold")
ax1.yaxis.grid(True, linestyle="--", alpha=0.6)
ax1.set_axisbelow(True)

# 成長幅（初日→最終日）
daily2 = df.groupby(["名前", "日付"])["スコア"].mean().reset_index()
growth = {}
for name, grp in daily2.groupby("名前"):
    grp = grp.sort_values("日付")
    growth[name] = grp["スコア"].iloc[-1] - grp["スコア"].iloc[0]

growth_s = pd.Series(growth).loc[names_sorted]
bar_colors = [COLORS.get(n) for n in growth_s.index]
bars = ax2.barh(growth_s.index[::-1], growth_s.values[::-1],
                color=bar_colors[::-1], alpha=0.8, edgecolor="white")

for bar, val in zip(bars, growth_s.values[::-1]):
    ax2.text(bar.get_width() + 0.1, bar.get_y() + bar.get_height() / 2,
             f"{val:+.1f}点", va="center", fontsize=11, fontweight="bold")

ax2.axvline(0, color="black", linewidth=1)
ax2.set_xlabel("スコア変化（初日→最終日）", fontsize=12)
ax2.set_title("成長幅ランキング", fontsize=14, fontweight="bold")
ax2.xaxis.grid(True, linestyle="--", alpha=0.6)
ax2.set_axisbelow(True)

plt.suptitle("受講者パフォーマンス詳細分析", fontsize=16, fontweight="bold", y=1.01)
plt.tight_layout()
plt.savefig("graph6_boxplot_growth.png", dpi=150, bbox_inches="tight")
plt.close()
print("graph6_boxplot_growth.png 保存完了")
