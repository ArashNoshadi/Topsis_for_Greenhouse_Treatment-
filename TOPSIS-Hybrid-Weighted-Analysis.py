"""
=============================================================
  TOPSIS Hybrid Weighted Analysis
  (AHP Subjective + Shannon Entropy Objective Weights)
=============================================================
"""

import sys
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import seaborn as sns

# ============================================================
#                   CONFIGURATION
#     Only edit this section for each new project
# ============================================================

IN_PATH    = r"C:\Users\a.noshadi\Input.xlsx"   # Path to input Excel file
SHEET_NAME = "Sheet1"        # Sheet name in the Excel file
OUT_DIR    = r"C:\Users\a.noshadi\\Topsis"       # Output directory

OUT_TXT   = "scored_results.txt"
OUT_EXCEL = "scored_results.xlsx"
OUT_PLOT  = "composite_scores.png"

# ============================================================


# ------------------------------------------------------------
#  Helper functions
# ------------------------------------------------------------

def read_structured_excel(path: str, sheet: str):
    """
    Reads Excel file with the structured format:
      Row 1 → Column names
      Row 2 → Weights
      Row 3 → Directions
      Row 4+ → Data
    Returns: col_names, weight_row, direction_row, data_df
    """
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


def parse_direction(raw_dir: str) -> str | None:
    """
    Parses direction string:
      '+' or 'positive' → '+'
      '-' or 'negative' → '-'
      'N/A', 'NA', 'none', '' → None (exclude trait)
    """
    d = str(raw_dir).strip().upper()
    if d in ("N/A", "NA", "NONE", "NAN", ""):
        return None
    if d.startswith("+") or d == "POSITIVE" or d == "P":
        return "+"
    if d.startswith("-") or d == "NEGATIVE" or d == "N":
        return "-"
    return None


def calculate_shannon_entropy_weights(X: np.ndarray) -> np.ndarray:
    """
    Objective weights via Shannon Entropy method.
    X: shape (n_samples, n_traits)
    Returns: weight array of length n_traits
    """
    X_pos = X.copy().astype(float)

    # Shift negatives so all values are positive
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

    d = np.maximum(1.0 - E, 0.0)   # divergence (≥ 0)
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

        # Parse direction
        dir_char = parse_direction(direction_row[i])
        if dir_char is None:
            print(f"  ↳ Trait '{name}': direction is N/A — excluded from analysis.")
            continue

        # Parse AHP weight
        try:
            w = float(weight_row[i])
            if np.isnan(w) or w < 0:
                raise ValueError
        except (ValueError, TypeError):
            print(f"  ↳ Trait '{name}': weight missing/invalid — set to 1.0")
            w = 1.0

        # Check column exists in data
        if name not in data_df.columns:
            print(f"  ↳ Trait '{name}': column not found in data — excluded.")
            continue

        traits.append(name)
        ahp_w_raw[name]  = w
        directions[name] = dir_char

    if not traits:
        print("[ERROR] No valid traits found. Check your Excel format.")
        sys.exit(1)

    print(f"\nTraits included in analysis ({len(traits)}): {traits}")

    # ── 4. Build decision matrix X ─────────────────────────
    n, m = len(data_df), len(traits)
    X = np.zeros((n, m))
    for j, t in enumerate(traits):
        X[:, j] = pd.to_numeric(data_df[t], errors="coerce").fillna(0).values

    # ── 5. AHP (subjective) weights ────────────────────────
    ahp_arr = np.array([ahp_w_raw[t] for t in traits], dtype=float)
    total_ahp = ahp_arr.sum()
    if total_ahp == 0:
        print("  ↳ All AHP weights are zero — using equal AHP weights.")
        ahp_arr = np.ones(m)
        total_ahp = float(m)
    ahp_normalized = ahp_arr / total_ahp   # normalized to sum = 1

    # ── 6. Shannon Entropy (objective) weights ─────────────
    shannon_arr = calculate_shannon_entropy_weights(X)

    # ── 7. Combined hybrid weights ─────────────────────────
    combined_raw = ahp_normalized * shannon_arr
    sum_combined = combined_raw.sum()
    if sum_combined == 0:
        weights_final = np.ones(m) / m
    else:
        weights_final = combined_raw / sum_combined

    # ── 8. TOPSIS ──────────────────────────────────────────
    # (a) Vector normalization
    col_norms = np.sqrt(np.sum(X ** 2, axis=0))
    col_norms[col_norms == 0] = 1.0
    R = X / col_norms

    # (b) Weighted normalized matrix
    V = R * weights_final

    # (c) Ideal solutions
    A_plus  = np.zeros(m)
    A_minus = np.zeros(m)
    for j, t in enumerate(traits):
        if directions[t] == "+":
            A_plus[j]  = np.max(V[:, j])
            A_minus[j] = np.min(V[:, j])
        else:                                  # '-'
            A_plus[j]  = np.min(V[:, j])
            A_minus[j] = np.max(V[:, j])

    # (d) Euclidean distances
    dist_plus  = np.sqrt(np.sum((V - A_plus)  ** 2, axis=1))
    dist_minus = np.sqrt(np.sum((V - A_minus) ** 2, axis=1))

    # (e) Relative closeness (composite score)
    total_dist = dist_plus + dist_minus
    total_dist[total_dist == 0] = 1.0
    Ci = dist_minus / total_dist

    # ── 9. Build output DataFrame ──────────────────────────
    out_df = data_df[[treat_col]].copy()
    out_df["composite_score"] = Ci
    out_df["rank"] = (
        pd.Series(Ci)
        .rank(method="min", ascending=False)
        .astype("Int64")
        .values
    )

    for j, t in enumerate(traits):
        out_df[f"{t}_norm"]     = R[:, j]
        out_df[f"{t}_weighted"] = V[:, j]

    # ── 10. Save outputs ───────────────────────────────────

    # Text file (tab-separated)
    txt_path = out_path / OUT_TXT
    out_df.to_csv(txt_path, sep="\t", index=False, na_rep="NA", float_format="%.6f")
    print(f"\nText file saved:  {txt_path}")

    # Excel with two sheets: Results + Weights
    excel_path = out_path / OUT_EXCEL
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="Results", index=False)

        weights_df = pd.DataFrame({
            "Trait":                  traits,
            "Direction":              [directions[t]       for t in traits],
            "AHP_Weight_Input":       [ahp_arr[j]          for j, t in enumerate(traits)],
            "AHP_Weight_Normalized":  [ahp_normalized[j]   for j, t in enumerate(traits)],
            "Shannon_Entropy_Weight": [shannon_arr[j]       for j, t in enumerate(traits)],
            "Final_Combined_Weight":  [weights_final[j]     for j, t in enumerate(traits)],
        })
        weights_df.to_excel(writer, sheet_name="Weights", index=False)

    print(f"Excel file saved: {excel_path}")

    # ── 11. Plot ───────────────────────────────────────────
    sns.set_style("whitegrid")
    sns.set_context("paper", font_scale=1.5)

    plot_df = out_df[[treat_col, "composite_score"]].sort_values(
        "composite_score", ascending=False
    )
    plot_df[treat_col] = plot_df[treat_col].astype(str)

    plt.figure(figsize=(10, max(6, len(plot_df) * 0.45)))
    ax = sns.barplot(
        x="composite_score", y=treat_col, data=plot_df,
        palette="Blues_d", orient="h"
    )

    for bar in ax.patches:
        w = bar.get_width()
        if not np.isnan(w):
            ax.text(
                w + 0.008,
                bar.get_y() + bar.get_height() / 2,
                f"{w:.3f}",
                ha="left", va="center", fontsize=11, color="#333333"
            )

    ax.set_xlabel("Composite Score  (AHP + Shannon Entropy)", fontsize=13)
    ax.set_ylabel("Treatment", fontsize=13)
    ax.set_title("TOPSIS – Hybrid Weighted Composite Scores", fontsize=15, pad=12)
    ax.set_xlim(0, 1.10)
    plt.tight_layout()

    plot_path = out_path / OUT_PLOT
    plt.savefig(plot_path, dpi=600, bbox_inches="tight")
    plt.close()
    print(f"Plot saved:       {plot_path}")

    # ── 12. Print summary ──────────────────────────────────
    summary = out_df.sort_values("composite_score", ascending=False)
    print("\n" + "=" * 55)
    print("  RESULTS  (sorted by composite score)")
    print("=" * 55)
    print(summary[[treat_col, "composite_score", "rank"]].to_string(index=False))

    print("\n" + "=" * 55)
    print("  WEIGHTS SUMMARY")
    print("=" * 55)
    print(weights_df.to_string(index=False))


if __name__ == "__main__":
    main()
