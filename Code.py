import os
from pathlib import Path
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
import sys
# ---------------------- Configuration (do not change this part) ----------------------
IN_PATH = r"G:\Paper\plants  extracts of some of Meloidogyne javanica Fish blood\Analysis\topsis\Input.xlsx"
SHEET_NAME = "Sheet3"
OUT_DIR = r"G:\Paper\plants  extracts of some of Meloidogyne javanica Fish blood\Analysis\topsis\v1"
# Output txt file name and plot name
OUT_TXT = "scored_results.txt"
OUT_EXCEL = "scored_results.xlsx"
OUT_PLOT = "composite_scores.png"
# Traits dictionary (based on your input)
# Structure for each trait: "TraitName": {"weight_col": "<weight-col-name>", "direction_col": "<direction-col-name>"}
# Trim keys to handle any extra spaces/newlines.
TRAITS_RAW = {
    "RF": {"weight_col": "RF-wgt", "direction_col": "RF-drc"},
    "Eggs": {"weight_col": "Eggs-wgt", "direction_col": "Eggs-drc"},
    "Egg Masses": {"weight_col": "Egg Masses-wgt", "direction_col": "Egg Masses-drc"},
    "Gall": {"weight_col": "Gall-wgt", "direction_col": "Gall-drc"},
    "SDW": {"weight_col": "SDW-wgt", "direction_col": "SDW-drc"},
    "SFW": {"weight_col": "SFW-wgt", "direction_col": "SFW-drc"},
    "RFW": {"weight_col": "RFW-wgt", "direction_col": "RFW-drc"},
    "SL": {"weight_col": "SL-wgt", "direction_col": "SL-drc"}
}
# ------------------------------------------------------------------------------------
def clean_traits(raw):
    """
    Convert the provided format to standard structure:
    { trait_name: {"value_col": trait_name, "weight_col": wcol, "direction_col": dcol} }
    and trim the keys.
    """
    out = {}
    for k, v in raw.items():
        trait = str(k).strip()
        weight_col = v.get("weight_col")
        direction_col = v.get("direction_col")
        if weight_col:
            weight_col = str(weight_col).strip()
        if direction_col:
            direction_col = str(direction_col).strip()
        out[trait] = {
            "value_col": trait,
            "weight_col": weight_col,
            "direction_col": direction_col
        }
    return out
def find_column_case_insensitive(df, name):
    """Return the actual column name in df via case-insensitive comparison or None"""
    if name is None:
        return None
    name_lower = str(name).strip().lower()
    for c in df.columns:
        if c.strip().lower() == name_lower:
            return c
    return None
def main():
    # Prepare output directory
    out_path = Path(OUT_DIR)
    out_path.mkdir(parents=True, exist_ok=True)
    # Read Excel
    try:
        df = pd.read_excel(IN_PATH, sheet_name=SHEET_NAME, engine='openpyxl')
    except Exception as e:
        print("Error reading Excel file:", e)
        sys.exit(1)
    # Find Treatment column (case-insensitive)
    treat_col = find_column_case_insensitive(df, "Treatment")
    if treat_col is None:
        print("Column 'Treatment' not found. Please ensure the Excel file has a 'Treatment' column.")
        sys.exit(1)
    # Clean and build standard TRAITS
    TRAITS = clean_traits(TRAITS_RAW)
    # Resolve columns (case-insensitive) and process
    resolved = {}
    valid_traits = []
    for trait, cfg in TRAITS.items():
        val_col = find_column_case_insensitive(df, cfg.get("value_col"))
        w_col = find_column_case_insensitive(df, cfg.get("weight_col"))
        d_col = find_column_case_insensitive(df, cfg.get("direction_col"))
        resolved[trait] = {"value_col": val_col, "weight_col": w_col, "direction_col": d_col}
        if val_col is not None:
            valid_traits.append(trait)
        else:
            print(f"Warning: Value column for trait '{trait}' (searching for '{cfg.get('value_col')}') not found; this trait will be ignored.")
    if len(valid_traits) == 0:
        print("No valid traits found — check that trait column names in Excel match TRAITS.")
        sys.exit(1)
    out_df = df.copy()
    # Calculate directions and weights
    weight_vals = {}
    direction_vals = {}
    for trait in valid_traits:
        info = resolved[trait]
        wcol = info["weight_col"]
        dcol = info["direction_col"]
        # Weight: if weight column exists, take mean (usually constant)
        w_val = None
        if wcol is not None and wcol in out_df.columns:
            wseries = pd.to_numeric(out_df[wcol], errors='coerce')
            if not wseries.isnull().all():
                w_val = wseries.mean(skipna=True)
        weight_vals[trait] = w_val
        # Direction: if no direction column -> default '+'
        dir_char = '+'
        if dcol is not None and dcol in out_df.columns:
            dirs = out_df[dcol].astype(str).str.strip().str[0].fillna('+')  # First character only
            uniq_dirs = dirs.unique()
            if len(uniq_dirs) == 1:
                dir_char = uniq_dirs[0]
            else:
                print(f"Warning: Mixed directions for '{trait}', defaulting to '+'")
        direction_vals[trait] = dir_char
    # Determine final weights: if no weights given -> equal weights
    given_weights = {t: w for t, w in weight_vals.items() if w is not None and not np.isnan(w)}
    weights_final = {}
    if not given_weights:
        equal_weight = 1.0 / len(valid_traits)
        for t in valid_traits:
            weights_final[t] = equal_weight
    else:
        total_weight = sum(given_weights.values())
        for t in valid_traits:
            w = weight_vals.get(t, 0.0)
            weights_final[t] = w / total_weight if total_weight > 0 else 1.0 / len(valid_traits)
    # Build decision matrix X
    n = len(out_df)
    m = len(valid_traits)
    traits = valid_traits
    X = np.zeros((n, m))
    for j, t in enumerate(traits):
        vcol = resolved[t]['value_col']
        X[:, j] = pd.to_numeric(out_df[vcol], errors='coerce').fillna(0)
    # TOPSIS normalization: vector normalization
    column_norms = np.sqrt(np.sum(X**2, axis=0))
    column_norms[column_norms == 0] = 1  # Avoid division by zero
    R = X / column_norms
    # Weighted normalized matrix V
    W = np.array([weights_final[t] for t in traits])
    V = R * W
    # Determine Positive Ideal Solution (PIS) and Negative Ideal Solution (NIS)
    A_plus = np.zeros(m)
    A_minus = np.zeros(m)
    for j, t in enumerate(traits):
        dir = direction_vals[t]
        if dir == '+':
            A_plus[j] = np.max(V[:, j])
            A_minus[j] = np.min(V[:, j])
        elif dir == '-':
            A_plus[j] = np.min(V[:, j])
            A_minus[j] = np.max(V[:, j])
    # Calculate distances
    dist_plus = np.sqrt(np.sum((V - A_plus)**2, axis=1))
    dist_minus = np.sqrt(np.sum((V - A_minus)**2, axis=1))
    # Relative closeness (composite score)
    total_dist = dist_plus + dist_minus
    total_dist[total_dist == 0] = 1  # Avoid division by zero
    Ci = dist_minus / total_dist
    out_df["composite_score"] = Ci
    # Add normalized and weighted to out_df
    for j, t in enumerate(traits):
        out_df[f"{t}_norm"] = R[:, j]
        out_df[f"{t}_weighted"] = V[:, j]
    # Ranking (rank 1 = best)
    out_df["rank"] = out_df["composite_score"].rank(method="min", ascending=False).astype('Int64')
    # Save text file (tab-separated) with main columns + normalized, weighted, composite + rank
    txt_cols = [treat_col, "composite_score", "rank"] + \
               [f"{t}_norm" for t in valid_traits] + [f"{t}_weighted" for t in valid_traits]
    txt_cols_available = [c for c in txt_cols if c in out_df.columns]
    txt_path = out_path / OUT_TXT
    out_df[txt_cols_available].to_csv(txt_path, sep='\t', index=False, na_rep="NA", float_format="%.6f")
    print(f"Output text file saved: {txt_path}")
    # Save to Excel
    excel_path = out_path / OUT_EXCEL
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        out_df[txt_cols_available].to_excel(writer, sheet_name='Results', index=False, na_rep="NA")
        # Save weights to a separate sheet
        weights_df = pd.DataFrame({
            'Trait': list(weights_final.keys()),
            'Weight': list(weights_final.values()),
            'Direction': [direction_vals[t] for t in weights_final.keys()]
        })
        weights_df.to_excel(writer, sheet_name='Weights', index=False)
    print(f"Output Excel file saved: {excel_path}")
    # Generate professional plot: use seaborn for better aesthetics
    sns.set_style("whitegrid")
    sns.set_context("paper", font_scale=1.5)  # For publication
    plot_df = out_df[[treat_col, "composite_score"]].copy()
    plot_df = plot_df.sort_values("composite_score", ascending=False)  # Sort descending to show best first
    plot_df[treat_col] = plot_df[treat_col].astype(str)  # Ensure string
    plt.figure(figsize=(10, 8))  # Adjusted size for better proportion
    ax = sns.barplot(x="composite_score", y=treat_col, data=plot_df, palette="Blues_d")  # Professional palette
    # Annotate values on bars
    for p in ax.patches:
        width = p.get_width()
        if not np.isnan(width):
            ax.text(width + 0.005, p.get_y() + p.get_height() / 2,
                    f"{width:.3f}", ha="left", va="center", fontsize=12, color="black")
    # Professional settings
    plt.xlabel("Composite Score", fontsize=14)
    plt.ylabel("Treatment", fontsize=14)
    plt.title("Composite Scores by Treatment (Sorted Descending)", fontsize=16)
    plt.xlim(0, 1.05)  # x range for annotation space
    plt.tight_layout()
    plot_path = out_path / OUT_PLOT
    plt.savefig(plot_path, dpi=600, bbox_inches='tight')  # Higher DPI for publication
    plt.close()
    print(f"Plot saved: {plot_path}")
    # Print summary of top 10 on console
    summary = out_df.sort_values("composite_score", ascending=False).head(10)
    display_cols = [treat_col, "composite_score", "rank"] + [f"{t}_norm" for t in valid_traits]
    display_cols = [c for c in display_cols if c in summary.columns]
    print("\n--- Top 10 (by composite_score) ---")
    print(summary[display_cols].to_string(index=False))
if __name__ == "__main__":
    main()