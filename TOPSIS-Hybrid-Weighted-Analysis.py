"""
=============================================================
  TOPSIS Hybrid Weighted Analysis
  (AHP Subjective + Shannon Entropy Objective Weights)
  -- Nature-Quality Plotting & Replicate Averaging Update --
  -- Added Target (N/A) Direction (Distance to Mean) --
=============================================================
"""

import sys
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.ticker as mticker
import matplotlib.colors as mcolors
from matplotlib.cm import ScalarMappable
import seaborn as sns

# ============================================================
#               CONFIGURATION
#     Only edit this section for each new project
# ============================================================

IN_PATH    = r"C:\Users\a.noshadi\4.xlsx"   # Path to input Excel file
SHEET_NAME = "Sheet1"        # Sheet name in the Excel file
OUT_DIR    = r"C:\Users\a.noshadi\4"       # Output directory

OUT_TXT   = "scored_results.txt"
OUT_EXCEL = "scored_results.xlsx"
OUT_PLOT  = "composite_scores.png"

OUT_PLOT_LOLLIPOP = "lollipop_scores_sorted.png"
OUT_PLOT_DONUT    = "weights_distribution.png"
OUT_PLOT_HEATMAP  = "topsis_heatmap.png"

alpha = 0.70  # (Subjective)
beta  = 0.30  # (Objective)
REFERENCE_TREATMENTS = ["Control"]
# ============================================================

# ------------------------------------------------------------
#  Helper functions
# ------------------------------------------------------------

def read_structured_excel(path: str, sheet: str):
    raw = pd.read_excel(path, sheet_name=sheet, header=None, engine="openpyxl")

    if len(raw) < 4:
        raise ValueError(
            "Excel file must have at least 4 rows:\n"
            "  Row 1: Column names\n"
            "  Row 2: Weights\n"
            "  Row 3: Directions\n"
            "  Row 4+: Data"
        )

    col_names     = raw.iloc[0].astype(str).str.strip().tolist()
    weight_row    = raw.iloc[1].tolist()
    direction_row = raw.iloc[2].astype(str).str.strip().tolist()

    data_df = raw.iloc[3:].copy()
    data_df.columns = col_names
    data_df = data_df.reset_index(drop=True)

    return col_names, weight_row, direction_row, data_df

def parse_direction(raw_dir: str) -> str | float | None:
    """
    Parses direction string:
      '+' or 'positive' → '+'
      '-' or 'negative' → '-'
      'N/A', 'NA', 'none', '' → 'TARGET' (Mean of references)
      Number (e.g., '3', '200') → Returns the float number (Fixed Target)
    """
    d = str(raw_dir).strip().upper()
    
    # 1. بررسی حالت N/A (وابسته به تیمار مرجع)
    if d in ("N/A", "NA", "NONE", "NAN", ""):
        return "TARGET"  
        
    # 2. بررسی حالت‌های مثبت و منفی
    if d.startswith("+") or d == "POSITIVE" or d == "P":
        return "+"
    if d.startswith("-") or d == "NEGATIVE" or d == "N":
        return "-"
        
    # 3. بررسی حالت جدید: آیا ورودی یک عدد ثابت (مثل 3 یا 200) است؟
    try:
        val = float(raw_dir)
        return val
    except ValueError:
        return None

def calculate_shannon_entropy_weights(X: np.ndarray) -> np.ndarray:
    X_pos = X.copy().astype(float)
    col_min = np.min(X_pos, axis=0)
    for j in range(X_pos.shape[1]):
        if col_min[j] < 0:
            X_pos[:, j] += abs(col_min[j]) + 1e-6
        elif col_min[j] == 0:
            X_pos[:, j] += 1e-6

    col_sums = np.sum(X_pos, axis=0)
    col_sums[col_sums == 0] = 1.0
    P = X_pos / col_sums

    n = P.shape[0]
    if n <= 1:
        return np.ones(X.shape[1]) / X.shape[1]

    k = 1.0 / np.log(n)
    P_log_P = P * np.log(P + 1e-12)
    E = -k * np.sum(P_log_P, axis=0)

    d = np.maximum(1.0 - E, 0.0)
    total_d = np.sum(d)

    if total_d == 0:
        return np.ones(X.shape[1]) / X.shape[1]
    return d / total_d

# ------------------------------------------------------------
#  Main analysis
# ------------------------------------------------------------

def main():
    out_path = Path(OUT_DIR)
    out_path.mkdir(parents=True, exist_ok=True)

    # ── 1. Read Excel ──────────────────────────────────────
    print("Reading Excel file …")
    try:
        col_names, weight_row, direction_row, data_df = read_structured_excel(IN_PATH, SHEET_NAME)
    except Exception as e:
        print(f"[ERROR] {e}")
        sys.exit(1)

    # ── 2. Identify Treatment column ───────────────────────
    treat_col = None
    treat_idx = None
    for i, name in enumerate(col_names):
        if name.strip().lower() == "treatment":
            treat_col = name
            treat_idx = i
            break

    if treat_col is None:
        print("[ERROR] Column 'Treatment' not found in Row 1 of the Excel file.")
        sys.exit(1)

    # ── 3. Build trait list ────────────────────────────────
    traits      = []
    ahp_w_raw   = {}
    directions  = {}

    for i, name in enumerate(col_names):
        if i == treat_idx:
            continue
        dir_char = parse_direction(direction_row[i])
        if dir_char is None:
            continue
        try:
            w = float(weight_row[i])
            if np.isnan(w) or w < 0:
                raise ValueError
        except (ValueError, TypeError):
            w = 1.0

        if name not in data_df.columns:
            continue

        traits.append(name)
        ahp_w_raw[name]  = w
        directions[name] = dir_char

    if not traits:
        print("[ERROR] No valid traits found. Check your Excel format.")
        sys.exit(1)

    print(f"\nTraits included in analysis ({len(traits)}): {traits}")

    # ── 3.5 Average Replicates ─────────────────────────────
    print("\nProcessing data: converting to numeric and averaging replicates...")
    data_df[treat_col] = data_df[treat_col].astype(str).str.strip()
    
    for t in traits:
        data_df[t] = pd.to_numeric(data_df[t], errors="coerce").fillna(0)
    
    original_len = len(data_df)
    data_df = data_df.groupby(treat_col, as_index=False)[traits].mean()
    new_len = len(data_df)
    print(f"Aggregated {original_len} rows into {new_len} unique treatments.")

    # ── 4. Build decision matrix X ─────────────────────────
    n, m = len(data_df), len(traits)
    X = np.zeros((n, m))
    for j, t in enumerate(traits):
        X[:, j] = data_df[t].values

    # اعمال منطق ویژگی‌های هدف‌گرا (N/A یا یک عدد ثابت)
    for j, t in enumerate(traits):
        dir_val = directions[t]
        
        # اگر جهت N/A باشد (TARGET) یا کاربر یک عدد ثابت (float) وارد کرده باشد
        if dir_val == "TARGET" or isinstance(dir_val, float):
            
            if dir_val == "TARGET":
                # حالت اول: محاسبه هدف بر اساس میانگین تیمار(های) مرجع
                if REFERENCE_TREATMENTS:
                    ref_data = data_df[data_df[treat_col].isin(REFERENCE_TREATMENTS)]
                    if len(ref_data) == 0:
                        T_j = np.mean(X[:, j])
                    else:
                        T_j = ref_data[t].mean()
                else:
                    T_j = np.mean(X[:, j])
            else:
                # حالت دوم: استفاده از عدد ثابتِ فیزیولوژیک به عنوان هدف (مثل 3 یا 200)
                T_j = dir_val
            
            # تبدیل داده‌ها به قدر مطلق فاصله از عدد هدف
            X[:, j] = np.abs(X[:, j] - T_j)
            
            # تغییر جهت به ویژگی منفی (-) تا هرچه فاصله از هدف کمتر باشد، امتیاز بهتر شود
            directions[t] = "-"
    
    # ── 5. AHP (subjective) weights ────────────────────────

    ahp_arr = np.array([ahp_w_raw[t] for t in traits], dtype=float)
    total_ahp = ahp_arr.sum() # <-- جمع کل وزن‌های کارشناس
    if total_ahp == 0:
        ahp_arr = np.ones(m)
        total_ahp = float(m)
    ahp_normalized = ahp_arr / total_ahp # <-- تقسیم هر وزن بر مجموع کل

    # ── 6. Shannon Entropy (objective) weights ─────────────
    shannon_arr = calculate_shannon_entropy_weights(X)

    # ── 7. Combined hybrid weights (Linear Combination) ────
    # اعمال ضرایب تاثیر: 70 درصد کارشناس (AHP) و 30 درصد آنتروپی شانون
    

    # ترکیب خطی وزن‌ها
    weights_final = (alpha * ahp_normalized) + (beta * shannon_arr)

    # نرمال‌سازی مجدد برای اطمینان از اینکه جمع کل دقیقاً برابر با 1 می‌شود
    sum_combined = weights_final.sum()
    if sum_combined == 0:
        weights_final = np.ones(m) / m
    else:
        weights_final = weights_final / sum_combined

    # ── 8. TOPSIS ──────────────────────────────────────────
    col_norms = np.sqrt(np.sum(X ** 2, axis=0))
    col_norms[col_norms == 0] = 1.0
    R = X / col_norms
    V = R * weights_final

    A_plus  = np.zeros(m)
    A_minus = np.zeros(m)
    for j, t in enumerate(traits):
        eff_dir = directions[t]  # <-- تغییر: استفاده مستقیم از دیکشنری directions
        if eff_dir == "+":
            A_plus[j]  = np.max(V[:, j])
            A_minus[j] = np.min(V[:, j])
        else:
            A_plus[j]  = np.min(V[:, j])
            A_minus[j] = np.max(V[:, j])

    dist_plus  = np.sqrt(np.sum((V - A_plus)  ** 2, axis=1))
    dist_minus = np.sqrt(np.sum((V - A_minus) ** 2, axis=1))

    total_dist = dist_plus + dist_minus
    total_dist[total_dist == 0] = 1.0
    Ci = dist_minus / total_dist

    # ── 9. Build output DataFrame ──────────────────────────
    out_df = data_df[[treat_col]].copy()
    out_df["composite_score"] = Ci
    out_df["rank"] = pd.Series(Ci).rank(method="min", ascending=False).astype("Int64").values

    # برای جلوگیری از هشدار PerformanceWarning، ستون‌ها را ابتدا در یک دیکشنری جمع می‌کنیم
    new_columns = {}
    for j, t in enumerate(traits):
        new_columns[f"{t}_norm"]     = R[:, j]
        new_columns[f"{t}_weighted"] = V[:, j]
        
    # همه ستون‌های جدید را یکجا به دیتافریم اصلی متصل می‌کنیم
    new_df = pd.DataFrame(new_columns, index=out_df.index)
    out_df = pd.concat([out_df, new_df], axis=1)


    ## ── 10. Save outputs ───────────────────────────────────
    txt_path = out_path / OUT_TXT
    out_df.to_csv(txt_path, sep="\t", index=False, na_rep="NA", float_format="%.6f")
    
    # --- اضافه کردن جدول نتایج رتبه‌بندی به انتهای فایل تکست ---
    summary = out_df.sort_values("composite_score", ascending=False)
    with open(txt_path, "a", encoding="utf-8") as f:
        f.write("\n\n" + "=" * 55 + "\n")
        f.write("  RESULTS  (sorted by composite score)\n")
        f.write("=" * 55 + "\n")
        f.write(summary[[treat_col, "composite_score", "rank"]].to_string(index=False))
        f.write("\n")
    # -----------------------------------------------------------

    excel_path = out_path / OUT_EXCEL
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="Results", index=False)
        weights_df = pd.DataFrame({
            "Trait":                  traits,
            "Direction":              [directions[t]       for t in traits],
            "AHP_Weight_Input":       [ahp_arr[j]          for j, t in enumerate(traits)],
            "AHP_Weight_Normalized":  [ahp_normalized[j]   for j, t in enumerate(traits)],
            "Shannon_Entropy_Weight": [shannon_arr[j]      for j, t in enumerate(traits)],
            "Final_Combined_Weight":  [weights_final[j]    for j, t in enumerate(traits)],
        })
        weights_df.to_excel(writer, sheet_name="Weights", index=False)

    # ====================================================================
    # ── Global Nature-quality plot settings ─────────────────────────────
    # ====================================================================
    plt.rcParams.update({
        "font.family":        "sans-serif",
        "font.sans-serif":    ["Arial", "Helvetica Neue", "Helvetica", "DejaVu Sans"],
        "font.size":          8,
        "axes.titlesize":     9,
        "axes.titleweight":   "normal",
        "axes.labelsize":     8,
        "xtick.labelsize":    7,
        "ytick.labelsize":    7,
        "axes.linewidth":     0.6,
        "xtick.major.width":  0.6,
        "ytick.major.width":  0.6,
        "xtick.major.size":   3,
        "ytick.major.size":   3,
        "figure.facecolor":   "white",
        "axes.facecolor":     "white",
        "savefig.dpi":        600,
        "pdf.fonttype":       42,
        "ps.fonttype":        42,
    })

    plot_df = out_df[[treat_col, "composite_score"]].sort_values(
        "composite_score", ascending=False
    ).reset_index(drop=True)
    plot_df[treat_col] = plot_df[treat_col].astype(str)

    n_rows = len(plot_df)
    nature_width = 7.2

    # ====================================================================
    # ── 11. PLOT 1 – Professional Horizontal Bar Chart ──────────────────
    # ====================================================================
    fig_h = max(3.5, n_rows * 0.35 + 1.0)
    fig, ax = plt.subplots(figsize=(nature_width, fig_h))

    bar_colors = sns.color_palette("Blues", n_rows + 4)[4:]
    bar_colors = bar_colors[::-1]

    bars = ax.barh(
        plot_df[treat_col],
        plot_df["composite_score"],
        color=bar_colors,
        edgecolor="black",
        linewidth=0.4,
        height=0.65,
    )

    x_max = plot_df["composite_score"].max()
    for bar in bars:
        w = bar.get_width()
        if not np.isnan(w):
            ax.text(
                w + (x_max * 0.015),
                bar.get_y() + bar.get_height() / 2,
                f"{w:.3f}",
                ha="left", va="center",
                fontsize=7, color="black"
            )

    ax.set_xlabel("Composite Score (AHP–Shannon Entropy Hybrid)", labelpad=6)
    ax.set_ylabel("Treatment", labelpad=6)
    ax.set_title("TOPSIS Hybrid Weighted Composite Scores", pad=8)
    ax.set_xlim(0, min(x_max * 1.15, 1.0))
    ax.xaxis.grid(True, linestyle="--", linewidth=0.4, color="#E0E0E0")
    ax.set_axisbelow(True)
    sns.despine(ax=ax, top=True, right=True)

    fig.tight_layout()
    plot_path = out_path / OUT_PLOT
    fig.savefig(plot_path, dpi=600, bbox_inches="tight")
    plt.close(fig)
    print(f"Plot saved:       {plot_path}")

    # ====================================================================
    # ── 11.1 PLOT 2 – Professional Sorted Lollipop Chart ────────────────
    # ====================================================================
    fig_h = max(3.5, n_rows * 0.35 + 1.0)
    fig, ax = plt.subplots(figsize=(nature_width, fig_h))

    scores = plot_df["composite_score"].values
    labels = plot_df[treat_col].tolist()
    y_pos  = np.arange(n_rows)

    cmap = plt.cm.viridis
    norm = mcolors.Normalize(vmin=scores.min(), vmax=scores.max())
    dot_colors = cmap(norm(scores))

    ax.hlines(y=y_pos, xmin=0, xmax=scores, color="gray", linewidth=1.0, zorder=1, alpha=0.6)
    ax.scatter(scores, y_pos, c=dot_colors, s=50, zorder=3, edgecolors="black", linewidths=0.4)

    for i, val in enumerate(scores):
        ax.text(val + (x_max * 0.015), i, f"{val:.3f}", va="center", ha="left", fontsize=7)

    sm = ScalarMappable(cmap=cmap, norm=norm)
    sm.set_array([])
    cbar = fig.colorbar(sm, ax=ax, shrink=0.6, pad=0.03, aspect=25)
    cbar.set_label("Composite Score", fontsize=7)
    cbar.ax.tick_params(labelsize=6)
    cbar.outline.set_linewidth(0.4)

    ax.set_yticks(y_pos)
    ax.set_yticklabels(labels)
    ax.set_xlim(0, min(x_max * 1.15, 1.0))
    ax.set_title("Ranked Composite Scores", pad=8)
    ax.set_xlabel("Composite Score", labelpad=6)
    ax.set_ylabel("Treatment", labelpad=6)
    ax.xaxis.grid(True, linestyle="--", linewidth=0.4, color="#E0E0E0")
    ax.set_axisbelow(True)
    sns.despine(ax=ax, top=True, right=True, left=True)
    ax.tick_params(axis="y", which="both", left=False)

    fig.tight_layout()
    plot_path_lol = out_path / OUT_PLOT_LOLLIPOP
    fig.savefig(plot_path_lol, dpi=600, bbox_inches="tight")
    plt.close(fig)
    print(f"Plot saved:       {plot_path_lol}")

    # ====================================================================
    # ── 11.2 PLOT 3 – Clean Weights Distribution Chart (Colors + Patterns)
    # ====================================================================
    n_traits = len(traits)
    
    # تنظیم داینامیک ابعاد
    legend_cols = 3 if n_traits > 9 else 2
    fig_h_donut = 4.5 + (n_traits / legend_cols) * 0.25
    fig, ax = plt.subplots(figsize=(6.5, fig_h_donut))
    
    # پالت رنگی
    donut_colors = sns.color_palette("husl", n_traits)

    wedges, _, autotexts = ax.pie(
        weights_final,
        autopct=lambda p: f"{p:.1f}%" if p >= 4.5 else "",
        startangle=90,
        colors=donut_colors,
        wedgeprops=dict(width=0.4, edgecolor="white", linewidth=0.8), # حاشیه سفید
        pctdistance=0.8,
    )
    plt.setp(autotexts, size=6.5, color="white", weight="bold")

    # --- بخش جدید: اضافه کردن الگوهای هاشور (Hatch Patterns) ---
    # لیستی از الگوهای استاندارد برای متمایز کردن رنگ‌های نزدیک به هم
    hatch_patterns = ['///', '...', '\\\\\\', 'xxx', '---', 'ooo', '+++', '|||', '***']
    
    for i, wedge in enumerate(wedges):
        # اختصاص یک الگو به هر قطعه به صورت دوره‌ای
        wedge.set_hatch(hatch_patterns[i % len(hatch_patterns)])
    # ------------------------------------------------------------

    ax.text(0, 0, "Weights", ha="center", va="center", fontsize=8, color="#333")

    # راهنما (Legend) به همراه نمایش الگوها
    ax.legend(
        wedges, traits,
        title="Traits",
        title_fontsize=8,
        loc="upper center",
        bbox_to_anchor=(0.5, -0.05),
        ncol=legend_cols,
        fontsize=7,
        frameon=False,
        columnspacing=1.2,
        handleheight=1.5, # افزایش ارتفاع آیکون‌های راهنما برای دیده شدن بهتر هاشورها
        handlelength=1.5
    )
    ax.set_title("Combined Weights Distribution", pad=12)

    fig.tight_layout()
    plot_path_donut = out_path / OUT_PLOT_DONUT
    fig.savefig(plot_path_donut, dpi=600, bbox_inches="tight")
    plt.close(fig)
    print(f"Plot saved:       {plot_path_donut}")

    # ====================================================================
    # ── 11.3 PLOT 4 – Heatmap of Weighted Normalized Matrix ─────────────
    # ====================================================================
    cell_w = max(0.8, 10.0 / max(m, 1))
    fig_w  = max(9.0, m * cell_w + 2.5)  
    
    cell_h = max(0.3, 5.0 / max(n, 1))
    fig_h  = max(3.5, n * cell_h + 1.5)
    
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))

    heatmap_data = pd.DataFrame(V, columns=traits, index=data_df[treat_col].astype(str))
    
    annot_fs = max(5, min(8, 100 / max(n * m, 1)))

    g = sns.heatmap(
        heatmap_data,
        annot=True,
        fmt=".3f",
        cmap="YlGnBu",
        linewidths=0.2,
        linecolor="white",
        cbar_kws={"label": "Weighted Value", "shrink": 0.8},
        annot_kws={"size": annot_fs},
        ax=ax
    )

    cbar_ax = g.collections[0].colorbar.ax
    cbar_ax.tick_params(labelsize=6)
    cbar_ax.set_ylabel("Weighted Value", fontsize=7)

    ax.set_title("Weighted Decision Matrix", pad=10)
    ax.set_ylabel("Treatment", labelpad=6)
    ax.set_xlabel("Traits", labelpad=6)
    
    ax.tick_params(axis="x", rotation=35, labelsize=7)
    ax.tick_params(axis="y", rotation=0, labelsize=7)

    fig.tight_layout()
    plot_path_heat = out_path / OUT_PLOT_HEATMAP
    fig.savefig(plot_path_heat, dpi=600, bbox_inches="tight")
    plt.close(fig)
    print(f"Plot saved:       {plot_path_heat}")

    # ── 12. Print summary ──────────────────────────────────
    summary = out_df.sort_values("composite_score", ascending=False)
    print("\n" + "=" * 55)
    print("  RESULTS  (sorted by composite score)")
    print("=" * 55)
    print(summary[[treat_col, "composite_score", "rank"]].to_string(index=False))

if __name__ == "__main__":
    main()
