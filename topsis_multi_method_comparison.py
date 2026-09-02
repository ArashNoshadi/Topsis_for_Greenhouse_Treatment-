"""
=================================================================================
  MULTI-METHOD MCDM DECISION-SUPPORT PIPELINE (TOPSIS / VIKOR / MARCOS /
  WASPAS / EDAS) FOR MULTI-CRITERIA TREATMENT RANKING
  (Extended version — original Hybrid AHP-Shannon Entropy TOPSIS preserved
   100% unchanged; 6 additional methods plus an expanded cross-method
   comparison module)
=================================================================================

METHODOLOGICAL RATIONALE (why each method was added)
--------------------------------------------------------------------------------
  Two independent axes are tested, so that the source of any disagreement
  between methods can be pinpointed rather than lumped together:

  AXIS A — effect of the WEIGHTING SCHEME (all three use TOPSIS aggregation):
  1) Hybrid AHP-Shannon Entropy TOPSIS  -> ORIGINAL METHOD, kept 100% unchanged.
  2) Simple (equal-weight) TOPSIS       -> classic textbook baseline with no
     subjective (AHP) or objective (Entropy/CRITIC) weighting at all. Needed
     as a neutral reference point to show what the weighting scheme actually
     contributes to the final ranking.
  3) Hybrid AHP-CRITIC TOPSIS           -> Shannon Entropy only looks at the
     dispersion of each trait in isolation; it is "correlation-blind". CRITIC
     (CRiteria Importance Through Intercriteria Correlation) additionally
     penalizes traits that are redundant (highly correlated) with other
     traits, so it does not over-weight two traits that carry almost the same
     biological information. Same AHP anchor (alpha/beta) as Method 1, so the
     two hybrids are directly comparable.

  AXIS B — effect of the AGGREGATION ALGORITHM (all four re-use the SAME
  Hybrid AHP-Entropy weights as Method 1, so only the ranking logic differs):
  4) VIKOR   -> TOPSIS (Euclidean distance to the ideal solution) is,
     empirically, the MCDM method most prone to rank reversal. VIKOR ranks by
     a compromise between group utility (S) and maximum individual regret (R).
  5) MARCOS  -> (Stevic et al., 2020) evaluates every alternative
     SIMULTANEOUSLY against an Ideal AND an Anti-Ideal reference point using
     ratio-scale normalization (not distance), combined into a single
     compromise utility f(K). Repeatedly reported in the recent MCDM
     literature as one of the most stable, rank-reversal-resistant methods
     available, with consistently strong rank correlation against other
     established methods.
  6) WASPAS  -> (Zavadskas et al., 2012) combines an additive Weighted Sum
     Model (WSM) and a multiplicative Weighted Product Model (WPM) into one
     score. In published rank-reversal stress-tests it has shown the LOWEST
     rank-reversal probability among common MCDM methods (lower than TOPSIS,
     VIKOR, COPRAS, SAW, EDAS), making it a very strong "quality control"
     cross-check on the final ranking.
  7) EDAS    -> (Keshavarz Ghorabaee et al., 2015) is the only method here
     that is NOT anchored to an extreme (ideal/anti-ideal) point: it measures
     distance from the AVERAGE solution instead. It is specifically reported
     to perform well when criteria strongly conflict with each other, which
     is common in agronomic/biological trait panels.

  Comparison module -> Spearman rank correlations, an overall Kendall's W
     concordance statistic, a discrimination-power metric (how decisively
     each method separates treatments), and a compact rank-heatmap across
     all 7 methods, summarized in Nature/Q1-style main-text and
     supplementary figures.

OUTPUT STRUCTURE (created under OUT_DIR)
--------------------------------------------------------------------------------
  01_Hybrid_AHP_Entropy_TOPSIS/   original method — unchanged outputs
  02_Simple_TOPSIS/               classic equal-weight TOPSIS
  03_CRITIC_AHP_TOPSIS/           AHP-CRITIC hybrid TOPSIS
  04_VIKOR/                       VIKOR compromise ranking
  05_MARCOS/                      MARCOS ideal/anti-ideal compromise utility
  06_WASPAS/                      WASPAS additive-multiplicative aggregation
  07_EDAS/                        EDAS distance-from-average ranking
  08_Method_Comparison/           cross-method comparison data + plots
      ├── main_text/              key figures for the manuscript body
      └── supplementary/          detailed diagnostic figures

  NOTE: the comparison folder moved from "05_..." to "08_..." because three
  new method folders (05-07) were inserted before it.
"""

import sys
from pathlib import Path

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
import matplotlib.colors as mcolors
from matplotlib.cm import ScalarMappable
import seaborn as sns

# ============================================================
#               CONFIGURATION
#     Only edit this section for each new project
# ============================================================

IN_PATH    = r"G:\Paper\plants  extracts of some of Meloidogyne javanica Fish blood\Analysis\topsis\Input1.xlsx"   # Path to input Excel file
SHEET_NAME = "Sheet3"                        # Sheet name in the Excel file
OUT_DIR    = r"G:\Paper\plants  extracts of some of Meloidogyne javanica Fish blood\Analysis\topsis\V4\Output"         # BASE output directory
                                              # (8 sub-folders are created inside it)

# Sub-folder names — no need to edit these
DIR_HYBRID     = "01_Hybrid_AHP_Entropy_TOPSIS"
DIR_SIMPLE     = "02_Simple_TOPSIS"
DIR_CRITIC     = "03_CRITIC_AHP_TOPSIS"
DIR_VIKOR      = "04_VIKOR"
DIR_MARCOS     = "05_MARCOS"
DIR_WASPAS     = "06_WASPAS"
DIR_EDAS       = "07_EDAS"
DIR_COMPARISON = "08_Method_Comparison"

alpha = 0.70  # AHP (subjective) share inside the two hybrid methods
beta  = 0.30  # Objective-weight (Entropy / CRITIC) share inside the two hybrids
REFERENCE_TREATMENTS = ["Control"]
VIKOR_V = 0.5        # VIKOR strategy weight: 0.5 = consensus/majority compromise,
                     # closer to 1 = "majority of criteria" strategy,
                     # closer to 0 = "individual regret / veto" strategy
WASPAS_LAMBDA = 0.5  # WASPAS blend between WSM (=1.0) and WPM (=0.0); 0.5 = equal
RATIO_EPSILON = 1e-6 # small offset so ratio-based methods (MARCOS, WASPAS) never
                     # divide by an exact zero (can occur in target-distance traits
                     # when a treatment exactly matches its target value)
# ============================================================

# ------------------------------------------------------------
#  Section 1 — Data reading & direction parsing (unchanged from original)
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


def parse_direction(raw_dir: str):
    """
    Parses direction string:
      '+' or 'positive' -> '+'
      '-' or 'negative' -> '-'
      'N/A', 'NA', 'none', '' -> 'TARGET' (distance to reference-treatment mean)
      Number (e.g., '3', '200') -> returns the float number (fixed target)
    """
    d = str(raw_dir).strip().upper()

    if d in ("N/A", "NA", "NONE", "NAN", ""):
        return "TARGET"

    if d.startswith("+") or d == "POSITIVE" or d == "P":
        return "+"
    if d.startswith("-") or d == "NEGATIVE" or d == "N":
        return "-"

    try:
        val = float(raw_dir)
        return val
    except ValueError:
        return None


def build_decision_matrix(data_df: pd.DataFrame, traits: list, directions_raw: dict, treat_col: str):
    """
    Builds the raw decision matrix X (n treatments x m traits) and resolves
    each trait's direction into a final '+' (benefit) or '-' (cost) code.
    'TARGET' directions (N/A) and fixed numeric targets are converted into an
    absolute distance-from-target column and re-coded as '-' (closer = better),
    exactly as in the original single-method script. This matrix and the
    resolved directions are shared by ALL four ranking methods below, so that
    differences between methods reflect only the weighting/aggregation logic,
    not differences in data preprocessing.
    """
    n, m = len(data_df), len(traits)
    X = np.zeros((n, m))
    for j, t in enumerate(traits):
        X[:, j] = data_df[t].values

    directions_resolved = dict(directions_raw)
    for j, t in enumerate(traits):
        dir_val = directions_raw[t]
        if dir_val == "TARGET" or isinstance(dir_val, float):
            if dir_val == "TARGET":
                if REFERENCE_TREATMENTS:
                    ref_data = data_df[data_df[treat_col].isin(REFERENCE_TREATMENTS)]
                    T_j = ref_data[t].mean() if len(ref_data) else np.mean(X[:, j])
                else:
                    T_j = np.mean(X[:, j])
            else:
                T_j = dir_val

            X[:, j] = np.abs(X[:, j] - T_j)
            directions_resolved[t] = "-"

    return X, directions_resolved


# ------------------------------------------------------------
#  Section 2 — Weighting schemes
# ------------------------------------------------------------

def calculate_shannon_entropy_weights(X: np.ndarray) -> np.ndarray:
    """Original Shannon Entropy objective weighting — unchanged."""
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


def calculate_critic_weights(X: np.ndarray, directions: list) -> np.ndarray:
    """
    CRITIC (CRiteria Importance Through Intercriteria Correlation).
    Combines contrast intensity (standard deviation, after min-max
    normalization respecting benefit/cost direction) with conflict
    (1 - Pearson correlation) between every pair of traits. A trait gets a
    higher weight when it varies a lot AND carries information that is not
    already captured by other traits — this is exactly the correlation
    sensitivity that Shannon Entropy lacks.
    """
    n, m = X.shape
    X_norm = np.zeros((n, m), dtype=float)
    for j in range(m):
        col = X[:, j]
        c_min, c_max = np.min(col), np.max(col)
        rng = c_max - c_min
        if rng == 0:
            X_norm[:, j] = 0.5
            continue
        if directions[j] == "+":
            X_norm[:, j] = (col - c_min) / rng
        else:
            X_norm[:, j] = (c_max - col) / rng

    std_j = np.std(X_norm, axis=0, ddof=1) if n > 1 else np.zeros(m)

    if m > 1 and n > 1:
        corr_matrix = np.corrcoef(X_norm, rowvar=False)
        corr_matrix = np.nan_to_num(corr_matrix, nan=0.0)
    else:
        corr_matrix = np.eye(m)

    conflict_j = np.sum(1 - corr_matrix, axis=1)
    C_j = std_j * conflict_j
    total_C = np.sum(C_j)

    if total_C == 0:
        return np.ones(m) / m
    return C_j / total_C


# ------------------------------------------------------------
#  Section 3 — Generic ranking engines (TOPSIS / VIKOR)
# ------------------------------------------------------------

def run_topsis_core(X: np.ndarray, weights: np.ndarray, directions: list):
    """
    Generic TOPSIS engine: vector normalization -> weighting -> ideal /
    anti-ideal solutions -> Euclidean distances -> closeness coefficient Ci
    (higher Ci = better). Shared by the Hybrid, Simple, and CRITIC methods —
    only the `weights` vector differs between them.
    """
    col_norms = np.sqrt(np.sum(X ** 2, axis=0))
    col_norms[col_norms == 0] = 1.0
    R = X / col_norms
    V = R * weights

    m = X.shape[1]
    A_plus = np.zeros(m)
    A_minus = np.zeros(m)
    for j in range(m):
        if directions[j] == "+":
            A_plus[j] = np.max(V[:, j])
            A_minus[j] = np.min(V[:, j])
        else:
            A_plus[j] = np.min(V[:, j])
            A_minus[j] = np.max(V[:, j])

    dist_plus = np.sqrt(np.sum((V - A_plus) ** 2, axis=1))
    dist_minus = np.sqrt(np.sum((V - A_minus) ** 2, axis=1))

    total_dist = dist_plus + dist_minus
    total_dist[total_dist == 0] = 1.0
    Ci = dist_minus / total_dist

    return Ci, R, V, dist_plus, dist_minus


def run_vikor_core(X: np.ndarray, weights: np.ndarray, directions: list, v: float = 0.5):
    """
    VIKOR compromise ranking method (Opricovic & Tzeng).
    S_i  = weighted sum of normalized regrets   (group utility, lower=better)
    R_i  = maximum weighted normalized regret   (individual regret, lower=better)
    Q_i  = v*(S normalized) + (1-v)*(R normalized)   (compromise index, lower=better)
    Provides an aggregation logic fundamentally different from TOPSIS'
    Euclidean distance-to-ideal, used here as an independent cross-check on
    ranking stability / susceptibility to rank reversal.
    """
    n, m = X.shape
    f_star = np.zeros(m)
    f_minus = np.zeros(m)
    for j in range(m):
        if directions[j] == "+":
            f_star[j] = np.max(X[:, j])
            f_minus[j] = np.min(X[:, j])
        else:
            f_star[j] = np.min(X[:, j])
            f_minus[j] = np.max(X[:, j])

    denom = f_star - f_minus
    denom[denom == 0] = 1e-9

    terms = weights[np.newaxis, :] * (f_star[np.newaxis, :] - X) / denom[np.newaxis, :]

    S = np.sum(terms, axis=1)
    R = np.max(terms, axis=1)

    S_star, S_minus = np.min(S), np.max(S)
    R_star, R_minus = np.min(R), np.max(R)
    S_range = (S_minus - S_star) if (S_minus - S_star) != 0 else 1e-9
    R_range = (R_minus - R_star) if (R_minus - R_star) != 0 else 1e-9

    Q = v * (S - S_star) / S_range + (1 - v) * (R - R_star) / R_range
    return S, R, Q, terms


def run_marcos_core(X: np.ndarray, weights: np.ndarray, directions: list, epsilon: float = RATIO_EPSILON):
    """
    MARCOS -- Measurement of Alternatives and Ranking according to
    COmpromise Solution (Stevic, Pamucar, Puska & Chatterjee, 2020).

    Unlike TOPSIS (which measures Euclidean DISTANCE to a single ideal
    point), MARCOS evaluates every alternative simultaneously against an
    Ideal (AI) and an Anti-Ideal (AAI) reference point using RATIO-scale
    normalization, then blends the two resulting utility degrees (K+, K-)
    into one compromise utility f(K) (higher = better). It is repeatedly
    reported as one of the most rank-reversal-resistant MCDM methods
    available, making it a strong independent cross-check on the ranking.

    A small epsilon is added to X because ratio-scale normalization is
    undefined at exactly zero (can occur in target-distance traits when a
    treatment matches its target exactly).
    """
    Xs = X + epsilon
    n, m = Xs.shape

    x_ai = np.zeros(m)
    x_aai = np.zeros(m)
    for j in range(m):
        if directions[j] == "+":
            x_ai[j] = np.max(Xs[:, j])
            x_aai[j] = np.min(Xs[:, j])
        else:
            x_ai[j] = np.min(Xs[:, j])
            x_aai[j] = np.max(Xs[:, j])
    x_ai[x_ai == 0] = epsilon
    x_aai[x_aai == 0] = epsilon

    N = np.zeros((n, m))
    n_ai = np.zeros(m)
    n_aai = np.zeros(m)
    for j in range(m):
        if directions[j] == "+":
            N[:, j] = Xs[:, j] / x_ai[j]
            n_ai[j] = 1.0
            n_aai[j] = x_aai[j] / x_ai[j]
        else:
            N[:, j] = x_ai[j] / Xs[:, j]
            n_ai[j] = 1.0
            n_aai[j] = x_ai[j] / x_aai[j]

    V = N * weights
    v_ai = n_ai * weights
    v_aai = n_aai * weights

    S = np.sum(V, axis=1)
    S_ai = np.sum(v_ai) if np.sum(v_ai) != 0 else 1e-9
    S_aai = np.sum(v_aai) if np.sum(v_aai) != 0 else 1e-9

    K_plus = S / S_ai
    K_minus = S / S_aai

    denom = K_plus + K_minus
    denom = np.where(denom == 0, 1e-9, denom)
    f_Kplus = K_minus / denom
    f_Kminus = K_plus / denom
    f_Kplus = np.where(f_Kplus == 0, 1e-9, f_Kplus)
    f_Kminus = np.where(f_Kminus == 0, 1e-9, f_Kminus)

    f_K = (K_plus + K_minus) / (
        1 + (1 - f_Kplus) / f_Kplus + (1 - f_Kminus) / f_Kminus
    )
    return f_K, K_plus, K_minus, V


def run_waspas_core(X: np.ndarray, weights: np.ndarray, directions: list,
                     lam: float = WASPAS_LAMBDA, epsilon: float = RATIO_EPSILON):
    """
    WASPAS -- Weighted Aggregated Sum Product Assessment
    (Zavadskas, Turskis, Antucheviciene & Zakarevicius, 2012).

    Combines two fundamentally different aggregation philosophies into one
    score: an additive Weighted Sum Model (Q1, compensatory: a bad trait can
    be offset by good traits) and a multiplicative Weighted Product Model
    (Q2, near non-compensatory: a very poor trait strongly penalizes the
    total regardless of other traits). Published rank-reversal stress-tests
    have found WASPAS to have the LOWEST rank-reversal probability among
    common MCDM methods, making it a strong stability cross-check.
    """
    Xs = X + epsilon
    n, m = Xs.shape
    R = np.zeros((n, m))
    for j in range(m):
        if directions[j] == "+":
            col_max = np.max(Xs[:, j])
            col_max = col_max if col_max != 0 else epsilon
            R[:, j] = Xs[:, j] / col_max
        else:
            col_min = np.min(Xs[:, j])
            col_min = col_min if col_min != 0 else epsilon
            R[:, j] = col_min / Xs[:, j]

    Q1 = np.sum(R * weights, axis=1)                 # Weighted Sum Model
    Q2 = np.prod(np.power(R, weights), axis=1)        # Weighted Product Model
    Q = lam * Q1 + (1 - lam) * Q2
    return Q, Q1, Q2, R


def run_edas_core(X: np.ndarray, weights: np.ndarray, directions: list):
    """
    EDAS -- Evaluation based on Distance from Average Solution
    (Keshavarz Ghorabaee, Zavadskas, Olfat & Turskis, 2015).

    The only method in this pipeline that is NOT anchored to an extreme
    (ideal/anti-ideal) reference point: every alternative is scored by how
    far it lies above (PDA, desirable) and below (NDA, undesirable) the
    AVERAGE solution across all treatments. Reported to perform
    particularly well when criteria strongly conflict with one another,
    which is common in multi-trait biological/agronomic panels.
    """
    n, m = X.shape
    AV = np.mean(X, axis=0)
    AV_safe = np.where(AV == 0, 1e-9, AV)

    PDA = np.zeros((n, m))
    NDA = np.zeros((n, m))
    for j in range(m):
        if directions[j] == "+":
            PDA[:, j] = np.maximum(0, X[:, j] - AV[j]) / AV_safe[j]
            NDA[:, j] = np.maximum(0, AV[j] - X[:, j]) / AV_safe[j]
        else:
            PDA[:, j] = np.maximum(0, AV[j] - X[:, j]) / AV_safe[j]
            NDA[:, j] = np.maximum(0, X[:, j] - AV[j]) / AV_safe[j]

    SP = np.sum(PDA * weights, axis=1)
    SN = np.sum(NDA * weights, axis=1)

    SP_max = np.max(SP) if np.max(SP) != 0 else 1.0
    SN_max = np.max(SN) if np.max(SN) != 0 else 1.0

    NSP = SP / SP_max
    NSN = 1 - (SN / SN_max)
    AS = (NSP + NSN) / 2
    return AS, SP, SN, PDA, NDA


# ------------------------------------------------------------
#  Section 4 — Cross-method comparison statistics
# ------------------------------------------------------------

def compute_discrimination_metrics(scores: np.ndarray) -> dict:
    """
    Quantifies how decisively a method differentiates between treatments.
    Computed on each method's normalized 0-1 "goodness" score so that
    methods with different native scales (TOPSIS Ci vs. VIKOR Q) are
    directly comparable.
    """
    s_sorted = np.sort(scores)[::-1]
    gaps = -np.diff(s_sorted)
    mean_val = np.mean(scores)
    return {
        "Range": float(np.max(scores) - np.min(scores)),
        "Std_Dev": float(np.std(scores, ddof=1)) if len(scores) > 1 else 0.0,
        "CV_percent": float(100 * np.std(scores, ddof=1) / mean_val) if (mean_val != 0 and len(scores) > 1) else 0.0,
        "Mean_Adjacent_Gap": float(np.mean(gaps)) if len(gaps) > 0 else 0.0,
    }


def kendalls_w(rank_matrix: np.ndarray) -> float:
    """
    Kendall's coefficient of concordance (W) across several rankings of the
    same alternatives (rank_matrix shape: n_methods x n_alternatives).
    0 = no agreement between methods, 1 = perfect agreement.
    (Tie-correction term omitted; acceptable for a supplementary diagnostic.)
    """
    k, n = rank_matrix.shape
    rank_sums = np.sum(rank_matrix, axis=0)
    mean_rank_sum = np.mean(rank_sums)
    S = np.sum((rank_sums - mean_rank_sum) ** 2)
    denom = (k ** 2) * (n ** 3 - n)
    if denom == 0:
        return 1.0
    W = 12 * S / denom
    return float(np.clip(W, 0, 1))


# ------------------------------------------------------------
#  Section 5 — Nature/Q1-style plotting theme & generic plot helpers
# ------------------------------------------------------------

def apply_nature_plot_style():
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


def plot_bar_scores(plot_df, treat_col, score_col, title, xlabel, out_file,
                     color_palette="Blues", higher_is_better=True):
    n_rows = len(plot_df)
    nature_width = 7.2
    fig_h = max(3.5, n_rows * 0.35 + 1.0)
    fig, ax = plt.subplots(figsize=(nature_width, fig_h))

    bar_colors = sns.color_palette(color_palette, n_rows + 4)[4:]
    if higher_is_better:
        bar_colors = bar_colors[::-1]

    bars = ax.barh(
        plot_df[treat_col], plot_df[score_col],
        color=bar_colors, edgecolor="black", linewidth=0.4, height=0.65,
    )

    x_max = plot_df[score_col].max()
    x_min = min(0, plot_df[score_col].min())
    label_offset = x_max * 0.015 if x_max else 0.01
    for bar in bars:
        w = bar.get_width()
        if not np.isnan(w):
            ax.text(w + label_offset, bar.get_y() + bar.get_height() / 2,
                     f"{w:.3f}", ha="left", va="center", fontsize=7, color="black")

    ax.set_xlabel(xlabel, labelpad=6)
    ax.set_ylabel("Treatment", labelpad=6)
    ax.set_title(title, pad=8)
    ax.set_xlim(x_min, x_max * 1.15 if x_max > 0 else 1.0)
    ax.xaxis.grid(True, linestyle="--", linewidth=0.4, color="#E0E0E0")
    ax.set_axisbelow(True)
    sns.despine(ax=ax, top=True, right=True)

    fig.tight_layout()
    fig.savefig(out_file, dpi=600, bbox_inches="tight")
    plt.close(fig)


def plot_lollipop(plot_df, treat_col, score_col, title, xlabel, out_file,
                   cmap_name="viridis"):
    n_rows = len(plot_df)
    nature_width = 7.2
    fig_h = max(3.5, n_rows * 0.35 + 1.0)
    fig, ax = plt.subplots(figsize=(nature_width, fig_h))

    scores = plot_df[score_col].values
    labels = plot_df[treat_col].tolist()
    y_pos = np.arange(n_rows)

    cmap = plt.get_cmap(cmap_name)
    norm = mcolors.Normalize(vmin=scores.min(), vmax=scores.max())
    dot_colors = cmap(norm(scores))

    ax.hlines(y=y_pos, xmin=0, xmax=scores, color="gray", linewidth=1.0, zorder=1, alpha=0.6)
    ax.scatter(scores, y_pos, c=dot_colors, s=50, zorder=3, edgecolors="black", linewidths=0.4)

    x_max = scores.max() if scores.max() != 0 else 1.0
    for i, val in enumerate(scores):
        ax.text(val + x_max * 0.015, i, f"{val:.3f}", va="center", ha="left", fontsize=7)

    sm = ScalarMappable(cmap=cmap, norm=norm)
    sm.set_array([])
    cbar = fig.colorbar(sm, ax=ax, shrink=0.6, pad=0.03, aspect=25)
    cbar.set_label(score_col, fontsize=7)
    cbar.ax.tick_params(labelsize=6)
    cbar.outline.set_linewidth(0.4)

    ax.set_yticks(y_pos)
    ax.set_yticklabels(labels)
    ax.set_xlim(min(0, scores.min()), x_max * 1.15)
    ax.set_title(title, pad=8)
    ax.set_xlabel(xlabel, labelpad=6)
    ax.set_ylabel("Treatment", labelpad=6)
    ax.xaxis.grid(True, linestyle="--", linewidth=0.4, color="#E0E0E0")
    ax.set_axisbelow(True)
    sns.despine(ax=ax, top=True, right=True, left=True)
    ax.tick_params(axis="y", which="both", left=False)

    fig.tight_layout()
    fig.savefig(out_file, dpi=600, bbox_inches="tight")
    plt.close(fig)


def plot_weights_donut(traits, weights_final, title, out_file):
    n_traits = len(traits)
    legend_cols = 3 if n_traits > 9 else 2
    fig_h_donut = 4.5 + (n_traits / legend_cols) * 0.25
    fig, ax = plt.subplots(figsize=(6.5, fig_h_donut))

    donut_colors = sns.color_palette("husl", n_traits)
    wedges, _, autotexts = ax.pie(
        weights_final,
        autopct=lambda p: f"{p:.1f}%" if p >= 4.5 else "",
        startangle=90,
        colors=donut_colors,
        wedgeprops=dict(width=0.4, edgecolor="white", linewidth=0.8),
        pctdistance=0.8,
    )
    plt.setp(autotexts, size=6.5, color="white", weight="bold")

    hatch_patterns = ['///', '...', '\\\\\\', 'xxx', '---', 'ooo', '+++', '|||', '***']
    for i, wedge in enumerate(wedges):
        wedge.set_hatch(hatch_patterns[i % len(hatch_patterns)])

    ax.text(0, 0, "Weights", ha="center", va="center", fontsize=8, color="#333")
    ax.legend(
        wedges, traits, title="Traits", title_fontsize=8, loc="upper center",
        bbox_to_anchor=(0.5, -0.05), ncol=legend_cols, fontsize=7, frameon=False,
        columnspacing=1.2, handleheight=1.5, handlelength=1.5,
    )
    ax.set_title(title, pad=12)

    fig.tight_layout()
    fig.savefig(out_file, dpi=600, bbox_inches="tight")
    plt.close(fig)


def plot_weighted_heatmap(V, traits, treat_labels, title, out_file):
    n, m = V.shape
    cell_w = max(0.8, 10.0 / max(m, 1))
    fig_w = max(9.0, m * cell_w + 2.5)
    cell_h = max(0.3, 5.0 / max(n, 1))
    fig_h = max(3.5, n * cell_h + 1.5)

    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    heatmap_data = pd.DataFrame(V, columns=traits, index=treat_labels)
    annot_fs = max(5, min(8, 100 / max(n * m, 1)))

    g = sns.heatmap(
        heatmap_data, annot=True, fmt=".3f", cmap="YlGnBu",
        linewidths=0.2, linecolor="white",
        cbar_kws={"label": "Weighted Value", "shrink": 0.8},
        annot_kws={"size": annot_fs}, ax=ax,
    )
    cbar_ax = g.collections[0].colorbar.ax
    cbar_ax.tick_params(labelsize=6)
    cbar_ax.set_ylabel("Weighted Value", fontsize=7)

    ax.set_title(title, pad=10)
    ax.set_ylabel("Treatment", labelpad=6)
    ax.set_xlabel("Traits", labelpad=6)
    ax.tick_params(axis="x", rotation=35, labelsize=7)
    ax.tick_params(axis="y", rotation=0, labelsize=7)

    fig.tight_layout()
    fig.savefig(out_file, dpi=600, bbox_inches="tight")
    plt.close(fig)


def plot_diagnostic_scatter(df, treat_col, x_col, y_col, color_col,
                             x_label, y_label, color_label, title, cmap_name, out_file):
    """
    Generic two-component diagnostic scatter (e.g. VIKOR's S vs R, MARCOS's
    K+ vs K-, WASPAS's Q1 vs Q2, EDAS's SP vs SN), colored by the method's
    own final compromise score. Reused across all four aggregation-focused
    methods so every method gets a visually consistent "how was this score
    built" figure, in addition to its ranked bar chart.
    """
    fig, ax = plt.subplots(figsize=(6.5, 5.5))
    sc = ax.scatter(
        df[x_col], df[y_col], c=df[color_col], cmap=cmap_name, s=90,
        edgecolors="black", linewidths=0.5, zorder=3,
    )
    for _, row in df.iterrows():
        ax.annotate(str(row[treat_col]), (row[x_col], row[y_col]), fontsize=6.5,
                     xytext=(4, 4), textcoords="offset points")

    cbar = fig.colorbar(sc, ax=ax, shrink=0.8)
    cbar.set_label(color_label, fontsize=8)
    ax.set_xlabel(x_label, labelpad=6)
    ax.set_ylabel(y_label, labelpad=6)
    ax.set_title(title, pad=10)
    ax.grid(True, linestyle="--", linewidth=0.4, color="#E0E0E0")
    ax.set_axisbelow(True)
    sns.despine(ax=ax, top=True, right=True)

    fig.tight_layout()
    fig.savefig(out_file, dpi=600, bbox_inches="tight")
    plt.close(fig)


def plot_rank_bump_chart(comp_df, treat_col, method_names, out_file):
    n = len(comp_df)
    n_methods = len(method_names)
    fig_h = max(4.0, n * 0.32 + 1.5)
    fig_w = max(7.2, n_methods * 1.35)
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    x_pos = np.arange(n_methods)
    palette = sns.color_palette("husl", n)

    order = comp_df.sort_values(f"{method_names[0]}_Rank")[treat_col].tolist()
    for idx, treat in enumerate(order):
        row = comp_df[comp_df[treat_col] == treat].iloc[0]
        y_vals = [row[f"{m}_Rank"] for m in method_names]
        ax.plot(x_pos, y_vals, marker="o", markersize=5, linewidth=1.3,
                 color=palette[idx], zorder=3, markeredgecolor="black", markeredgewidth=0.4)
        ax.text(x_pos[-1] + 0.10, y_vals[-1], str(treat), fontsize=6.5, va="center")

    ax.set_xticks(x_pos)
    ax.set_xticklabels([m.replace("_", " ") for m in method_names], fontsize=7, rotation=15, ha="right")
    ax.set_xlim(x_pos[0] - 0.35, x_pos[-1] + 1.35)
    max_rank = int(comp_df[[f"{m}_Rank" for m in method_names]].values.max())
    ax.set_yticks(np.arange(1, max_rank + 1))
    ax.invert_yaxis()
    ax.set_ylabel("Rank (1 = best)", labelpad=6)
    ax.set_title("Treatment Rank Consistency Across MCDM Methods", pad=10)
    ax.grid(True, axis="y", linestyle="--", linewidth=0.4, color="#E0E0E0")
    ax.set_axisbelow(True)
    sns.despine(ax=ax, top=True, right=True)

    fig.tight_layout()
    fig.savefig(out_file, dpi=600, bbox_inches="tight")
    plt.close(fig)


def plot_goodness_grouped_bar(comp_df, treat_col, method_names, out_file):
    n = len(comp_df)
    n_methods = len(method_names)
    df_sorted = comp_df.sort_values(f"{method_names[0]}_Goodness", ascending=False)
    x = np.arange(n)
    width = 0.82 / n_methods
    fig_w = max(7.5, n * 0.7 + n_methods * 0.3)
    fig, ax = plt.subplots(figsize=(fig_w, 4.8))
    colors = sns.color_palette("tab10", n_methods)

    for i, m in enumerate(method_names):
        ax.bar(x + i * width, df_sorted[f"{m}_Goodness"].values, width=width,
                label=m.replace("_", " "), color=colors[i], edgecolor="black", linewidth=0.3)

    ax.set_xticks(x + width * (n_methods - 1) / 2)
    ax.set_xticklabels(df_sorted[treat_col].astype(str).tolist(), rotation=35, ha="right", fontsize=7.5)
    ax.set_ylabel("Normalized Goodness Score (0–1, higher = better)", labelpad=6)
    ax.set_title("Comparative Composite Scores Across Methods", pad=10)
    ax.legend(fontsize=6.8, frameon=False, ncol=min(4, n_methods), loc="upper center",
              bbox_to_anchor=(0.5, -0.18))
    ax.yaxis.grid(True, linestyle="--", linewidth=0.4, color="#E0E0E0")
    ax.set_axisbelow(True)
    sns.despine(ax=ax, top=True, right=True)

    fig.tight_layout()
    fig.savefig(out_file, dpi=600, bbox_inches="tight")
    plt.close(fig)


def plot_spearman_heatmap(corr_df, out_file):
    k = len(corr_df.columns)
    fig_w = max(5.5, k * 0.95)
    fig_h = max(4.8, k * 0.85)
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))
    labels = [c.replace("_", " ") for c in corr_df.columns]
    sns.heatmap(
        corr_df, annot=True, fmt=".3f", cmap="RdYlBu_r", vmin=-1, vmax=1,
        square=True, linewidths=0.4, linecolor="white",
        cbar_kws={"label": "Spearman \u03c1", "shrink": 0.8},
        xticklabels=labels, yticklabels=labels, ax=ax, annot_kws={"size": 7.5},
    )
    ax.set_title("Spearman Rank Correlation Between Methods", pad=10)
    ax.tick_params(axis="x", rotation=35, labelsize=7)
    ax.tick_params(axis="y", rotation=0, labelsize=7)

    fig.tight_layout()
    fig.savefig(out_file, dpi=600, bbox_inches="tight")
    plt.close(fig)


def plot_discrimination_bar(disc_df, out_file):
    fig_w = max(6.0, len(disc_df) * 1.05)
    fig, ax = plt.subplots(figsize=(fig_w, 4.4))
    labels = disc_df["Method"].str.replace("_", " ")
    colors = sns.color_palette("magma", len(disc_df))
    bars = ax.bar(labels, disc_df["CV_percent"], color=colors, edgecolor="black", linewidth=0.4)

    for bar, val in zip(bars, disc_df["CV_percent"]):
        ax.text(bar.get_x() + bar.get_width() / 2, bar.get_height() + 0.5,
                 f"{val:.1f}%", ha="center", va="bottom", fontsize=7.5)

    ax.set_ylabel("Coefficient of Variation (%) of Goodness Scores", labelpad=6)
    ax.set_title("Discrimination Power: Treatment-Separation Ability by Method", pad=10)
    plt.setp(ax.get_xticklabels(), rotation=25, ha="right", fontsize=8)
    ax.yaxis.grid(True, linestyle="--", linewidth=0.4, color="#E0E0E0")
    ax.set_axisbelow(True)
    sns.despine(ax=ax, top=True, right=True)

    fig.tight_layout()
    fig.savefig(out_file, dpi=600, bbox_inches="tight")
    plt.close(fig)


def plot_rank_heatmap(comp_df, treat_col, method_names, out_file):
    """
    Compact treatments x methods rank matrix -- with 7 methods this gives a
    faster at-a-glance consistency check than the bump chart alone
    (green = consistently top-ranked, red = consistently bottom-ranked).
    """
    rank_cols = [f"{m}_Rank" for m in method_names]
    data = comp_df.set_index(treat_col)[rank_cols].copy()
    data.columns = [m.replace("_", " ") for m in method_names]

    n_treat, n_methods = data.shape
    fig_w = max(6.5, n_methods * 1.1 + 2.0)
    fig_h = max(3.5, n_treat * 0.42 + 1.5)
    fig, ax = plt.subplots(figsize=(fig_w, fig_h))

    sns.heatmap(
        data, annot=True, fmt="d", cmap="RdYlGn_r", linewidths=0.4, linecolor="white",
        cbar_kws={"label": "Rank (1 = best)", "shrink": 0.8}, ax=ax, annot_kws={"size": 8},
    )
    ax.set_title("Treatment Ranks Across All MCDM Methods", pad=10)
    ax.set_xlabel("Method", labelpad=6)
    ax.set_ylabel("Treatment", labelpad=6)
    ax.tick_params(axis="x", rotation=30, labelsize=7.5)
    ax.tick_params(axis="y", rotation=0, labelsize=7.5)

    fig.tight_layout()
    fig.savefig(out_file, dpi=600, bbox_inches="tight")
    plt.close(fig)


# ------------------------------------------------------------
#  Section 6 — Shared output writers
# ------------------------------------------------------------

def build_result_dataframe(data_df, treat_col, traits, Ci, R, V):
    out_df = data_df[[treat_col]].copy()
    out_df["composite_score"] = Ci
    out_df["rank"] = pd.Series(Ci).rank(method="min", ascending=False).astype("Int64").values

    new_columns = {}
    for j, t in enumerate(traits):
        new_columns[f"{t}_norm"] = R[:, j]
        new_columns[f"{t}_weighted"] = V[:, j]
    new_df = pd.DataFrame(new_columns, index=out_df.index)
    out_df = pd.concat([out_df, new_df], axis=1)
    return out_df


def save_topsis_txt_excel(out_df, treat_col, out_dir, txt_name, excel_name, weights_table=None):
    txt_path = out_dir / txt_name
    out_df.to_csv(txt_path, sep="\t", index=False, na_rep="NA", float_format="%.6f")

    summary = out_df.sort_values("composite_score", ascending=False)
    with open(txt_path, "a", encoding="utf-8") as f:
        f.write("\n\n" + "=" * 55 + "\n")
        f.write("  RESULTS  (sorted by composite score)\n")
        f.write("=" * 55 + "\n")
        f.write(summary[[treat_col, "composite_score", "rank"]].to_string(index=False))
        f.write("\n")

    excel_path = out_dir / excel_name
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="Results", index=False)
        if weights_table is not None:
            weights_table.to_excel(writer, sheet_name="Weights", index=False)

    return txt_path, excel_path


# ------------------------------------------------------------
#  Section 7 — Per-method runners
# ------------------------------------------------------------

def run_hybrid_method(data_df, treat_col, traits, X, dir_list, ahp_normalized, shannon_arr, out_dir):
    """Method 1 — ORIGINAL hybrid AHP-Shannon Entropy TOPSIS (unchanged math)."""
    weights_final = (alpha * ahp_normalized) + (beta * shannon_arr)
    s = weights_final.sum()
    weights_final = weights_final / s if s != 0 else np.ones(len(traits)) / len(traits)

    Ci, R, V, _, _ = run_topsis_core(X, weights_final, dir_list)
    out_df = build_result_dataframe(data_df, treat_col, traits, Ci, R, V)

    weights_table = pd.DataFrame({
        "Trait": traits,
        "Direction": dir_list,
        "AHP_Weight_Normalized": ahp_normalized,
        "Shannon_Entropy_Weight": shannon_arr,
        "Final_Combined_Weight": weights_final,
    })
    save_topsis_txt_excel(out_df, treat_col, out_dir, "scored_results.txt", "scored_results.xlsx", weights_table)

    plot_df = out_df[[treat_col, "composite_score"]].sort_values("composite_score", ascending=False).reset_index(drop=True)
    plot_df[treat_col] = plot_df[treat_col].astype(str)

    plot_bar_scores(plot_df, treat_col, "composite_score",
                     "TOPSIS Hybrid Weighted Composite Scores",
                     "Composite Score (AHP\u2013Shannon Entropy Hybrid)",
                     out_dir / "composite_scores.png")
    plot_lollipop(plot_df, treat_col, "composite_score",
                  "Ranked Composite Scores (Hybrid AHP\u2013Entropy)",
                  "Composite Score", out_dir / "lollipop_scores_sorted.png")
    plot_weights_donut(traits, weights_final, "Combined Weights Distribution (AHP\u2013Entropy)",
                        out_dir / "weights_distribution.png")
    plot_weighted_heatmap(V, traits, data_df[treat_col].astype(str),
                           "Weighted Decision Matrix (Hybrid AHP\u2013Entropy)",
                           out_dir / "topsis_heatmap.png")

    print(f"  -> Hybrid method outputs saved to: {out_dir}")
    return {
        "method": "Hybrid_AHP_Entropy", "treat_col": treat_col, "scores": Ci,
        "treatments": data_df[treat_col].astype(str).values,
        "higher_is_better": True, "weights": weights_final,
    }


def run_simple_method(data_df, treat_col, traits, X, dir_list, out_dir):
    """Method 2 — Simple / classic TOPSIS with equal criteria weights (baseline)."""
    m = len(traits)
    weights_equal = np.ones(m) / m

    Ci, R, V, _, _ = run_topsis_core(X, weights_equal, dir_list)
    out_df = build_result_dataframe(data_df, treat_col, traits, Ci, R, V)

    weights_table = pd.DataFrame({"Trait": traits, "Direction": dir_list, "Equal_Weight": weights_equal})
    save_topsis_txt_excel(out_df, treat_col, out_dir, "scored_results_simple.txt",
                           "scored_results_simple.xlsx", weights_table)

    plot_df = out_df[[treat_col, "composite_score"]].sort_values("composite_score", ascending=False).reset_index(drop=True)
    plot_df[treat_col] = plot_df[treat_col].astype(str)

    plot_bar_scores(plot_df, treat_col, "composite_score",
                     "Simple (Classic, Equal-Weight) TOPSIS Composite Scores",
                     "Composite Score (Equal Weights)",
                     out_dir / "composite_scores_simple.png", color_palette="Greys")
    plot_lollipop(plot_df, treat_col, "composite_score",
                  "Ranked Composite Scores (Simple TOPSIS)",
                  "Composite Score", out_dir / "lollipop_scores_simple.png", cmap_name="cividis")
    plot_weighted_heatmap(V, traits, data_df[treat_col].astype(str),
                           "Weighted Decision Matrix (Simple / Equal Weights)",
                           out_dir / "topsis_heatmap_simple.png")

    print(f"  -> Simple TOPSIS outputs saved to: {out_dir}")
    return {
        "method": "Simple_TOPSIS", "treat_col": treat_col, "scores": Ci,
        "treatments": data_df[treat_col].astype(str).values,
        "higher_is_better": True, "weights": weights_equal,
    }


def run_critic_method(data_df, treat_col, traits, X, dir_list, ahp_normalized, out_dir):
    """Method 3 — Hybrid AHP-CRITIC TOPSIS (correlation-aware objective weighting)."""
    critic_arr = calculate_critic_weights(X, dir_list)
    weights_final = (alpha * ahp_normalized) + (beta * critic_arr)
    s = weights_final.sum()
    weights_final = weights_final / s if s != 0 else np.ones(len(traits)) / len(traits)

    Ci, R, V, _, _ = run_topsis_core(X, weights_final, dir_list)
    out_df = build_result_dataframe(data_df, treat_col, traits, Ci, R, V)

    weights_table = pd.DataFrame({
        "Trait": traits,
        "Direction": dir_list,
        "AHP_Weight_Normalized": ahp_normalized,
        "CRITIC_Weight": critic_arr,
        "Final_Combined_Weight": weights_final,
    })
    save_topsis_txt_excel(out_df, treat_col, out_dir, "scored_results_critic.txt",
                           "scored_results_critic.xlsx", weights_table)

    plot_df = out_df[[treat_col, "composite_score"]].sort_values("composite_score", ascending=False).reset_index(drop=True)
    plot_df[treat_col] = plot_df[treat_col].astype(str)

    plot_bar_scores(plot_df, treat_col, "composite_score",
                     "TOPSIS Hybrid Weighted Composite Scores (AHP\u2013CRITIC)",
                     "Composite Score (AHP\u2013CRITIC Hybrid)",
                     out_dir / "composite_scores_critic.png", color_palette="Greens")
    plot_lollipop(plot_df, treat_col, "composite_score",
                  "Ranked Composite Scores (Hybrid AHP\u2013CRITIC)",
                  "Composite Score", out_dir / "lollipop_scores_critic.png", cmap_name="plasma")
    plot_weights_donut(traits, weights_final, "Combined Weights Distribution (AHP\u2013CRITIC)",
                        out_dir / "weights_distribution_critic.png")
    plot_weighted_heatmap(V, traits, data_df[treat_col].astype(str),
                           "Weighted Decision Matrix (Hybrid AHP\u2013CRITIC)",
                           out_dir / "topsis_heatmap_critic.png")

    print(f"  -> AHP-CRITIC TOPSIS outputs saved to: {out_dir}")
    return {
        "method": "CRITIC_AHP_TOPSIS", "treat_col": treat_col, "scores": Ci,
        "treatments": data_df[treat_col].astype(str).values,
        "higher_is_better": True, "weights": weights_final,
    }


def run_vikor_method(data_df, treat_col, traits, X, dir_list, weights_for_vikor, out_dir):
    """Method 4 — VIKOR compromise ranking (same weights as Method 1, different aggregation)."""
    S, R_, Q, terms = run_vikor_core(X, weights_for_vikor, dir_list, v=VIKOR_V)

    out_df = data_df[[treat_col]].copy()
    out_df["S"] = S
    out_df["R"] = R_
    out_df["Q"] = Q
    out_df["rank"] = pd.Series(Q).rank(method="min", ascending=True).astype("Int64").values

    txt_path = out_dir / "scored_results_vikor.txt"
    out_df.to_csv(txt_path, sep="\t", index=False, na_rep="NA", float_format="%.6f")
    summary = out_df.sort_values("Q", ascending=True)
    with open(txt_path, "a", encoding="utf-8") as f:
        f.write("\n\n" + "=" * 55 + "\n")
        f.write("  RESULTS  (sorted by VIKOR Q -- LOWER is better)\n")
        f.write("=" * 55 + "\n")
        f.write(summary[[treat_col, "S", "R", "Q", "rank"]].to_string(index=False))
        f.write("\n")

    excel_path = out_dir / "scored_results_vikor.xlsx"
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="Results", index=False)
        pd.DataFrame({"Trait": traits, "Direction": dir_list, "Weight_Used": weights_for_vikor}).to_excel(
            writer, sheet_name="Weights", index=False)

    out_df_str = out_df.copy()
    out_df_str[treat_col] = out_df_str[treat_col].astype(str)

    plot_df = out_df_str[[treat_col, "Q"]].sort_values("Q", ascending=True).reset_index(drop=True)
    plot_bar_scores(plot_df, treat_col, "Q", "VIKOR Compromise Ranking (Q Index)",
                     "VIKOR Q Index (lower = better)", out_dir / "vikor_compromise_scores.png",
                     color_palette="Oranges", higher_is_better=False)
    plot_diagnostic_scatter(out_df_str, treat_col, "S", "R", "Q",
                             "S \u2014 group utility (average weighted regret)",
                             "R \u2014 individual regret (maximum weighted regret)",
                             "Q (compromise index, lower = better)",
                             "VIKOR: Group Utility vs. Individual Regret",
                             "RdYlGn_r", out_dir / "vikor_SRQ_plot.png")
    plot_weighted_heatmap(terms, traits, data_df[treat_col].astype(str),
                           "Weighted Regret Contribution Matrix (VIKOR)", out_dir / "vikor_heatmap.png")

    print(f"  -> VIKOR outputs saved to: {out_dir}")
    return {
        "method": "VIKOR", "treat_col": treat_col, "scores": Q,
        "treatments": data_df[treat_col].astype(str).values,
        "higher_is_better": False, "weights": weights_for_vikor,
    }


def run_marcos_method(data_df, treat_col, traits, X, dir_list, weights_for_marcos, out_dir):
    """Method 5 — MARCOS compromise utility (same weights as Method 1)."""
    f_K, K_plus, K_minus, V = run_marcos_core(X, weights_for_marcos, dir_list)

    out_df = data_df[[treat_col]].copy()
    out_df["K_plus"] = K_plus
    out_df["K_minus"] = K_minus
    out_df["Utility_fK"] = f_K
    out_df["rank"] = pd.Series(f_K).rank(method="min", ascending=False).astype("Int64").values

    txt_path = out_dir / "scored_results_marcos.txt"
    out_df.to_csv(txt_path, sep="\t", index=False, na_rep="NA", float_format="%.6f")
    summary = out_df.sort_values("Utility_fK", ascending=False)
    with open(txt_path, "a", encoding="utf-8") as f:
        f.write("\n\n" + "=" * 55 + "\n")
        f.write("  RESULTS  (sorted by MARCOS utility f(K) -- HIGHER is better)\n")
        f.write("=" * 55 + "\n")
        f.write(summary[[treat_col, "K_plus", "K_minus", "Utility_fK", "rank"]].to_string(index=False))
        f.write("\n")

    excel_path = out_dir / "scored_results_marcos.xlsx"
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="Results", index=False)
        pd.DataFrame({"Trait": traits, "Direction": dir_list, "Weight_Used": weights_for_marcos}).to_excel(
            writer, sheet_name="Weights", index=False)

    out_df_str = out_df.copy()
    out_df_str[treat_col] = out_df_str[treat_col].astype(str)

    plot_df = out_df_str[[treat_col, "Utility_fK"]].sort_values("Utility_fK", ascending=False).reset_index(drop=True)
    plot_bar_scores(plot_df, treat_col, "Utility_fK", "MARCOS Compromise Utility Ranking",
                     "MARCOS Utility f(K) (higher = better)", out_dir / "marcos_utility_scores.png",
                     color_palette="Purples")
    plot_diagnostic_scatter(out_df_str, treat_col, "K_plus", "K_minus", "Utility_fK",
                             "K\u207a \u2014 utility relative to the Ideal Solution",
                             "K\u207b \u2014 utility relative to the Anti-Ideal Solution",
                             "f(K) utility (higher = better)",
                             "MARCOS: Utility Relative to Ideal vs. Anti-Ideal",
                             "viridis", out_dir / "marcos_KplusKminus_scatter.png")
    plot_weighted_heatmap(V, traits, data_df[treat_col].astype(str),
                           "Weighted Normalized Matrix (MARCOS)", out_dir / "marcos_heatmap.png")

    print(f"  -> MARCOS outputs saved to: {out_dir}")
    return {
        "method": "MARCOS", "treat_col": treat_col, "scores": f_K,
        "treatments": data_df[treat_col].astype(str).values,
        "higher_is_better": True, "weights": weights_for_marcos,
    }


def run_waspas_method(data_df, treat_col, traits, X, dir_list, weights_for_waspas, out_dir):
    """Method 6 — WASPAS additive-multiplicative aggregation (same weights as Method 1)."""
    Q, Q1, Q2, R = run_waspas_core(X, weights_for_waspas, dir_list, lam=WASPAS_LAMBDA)

    out_df = data_df[[treat_col]].copy()
    out_df["WSM_Q1"] = Q1
    out_df["WPM_Q2"] = Q2
    out_df["WASPAS_Q"] = Q
    out_df["rank"] = pd.Series(Q).rank(method="min", ascending=False).astype("Int64").values

    txt_path = out_dir / "scored_results_waspas.txt"
    out_df.to_csv(txt_path, sep="\t", index=False, na_rep="NA", float_format="%.6f")
    summary = out_df.sort_values("WASPAS_Q", ascending=False)
    with open(txt_path, "a", encoding="utf-8") as f:
        f.write("\n\n" + "=" * 55 + "\n")
        f.write("  RESULTS  (sorted by WASPAS Q -- HIGHER is better)\n")
        f.write("=" * 55 + "\n")
        f.write(summary[[treat_col, "WSM_Q1", "WPM_Q2", "WASPAS_Q", "rank"]].to_string(index=False))
        f.write("\n")

    excel_path = out_dir / "scored_results_waspas.xlsx"
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="Results", index=False)
        pd.DataFrame({"Trait": traits, "Direction": dir_list, "Weight_Used": weights_for_waspas}).to_excel(
            writer, sheet_name="Weights", index=False)

    out_df_str = out_df.copy()
    out_df_str[treat_col] = out_df_str[treat_col].astype(str)

    plot_df = out_df_str[[treat_col, "WASPAS_Q"]].sort_values("WASPAS_Q", ascending=False).reset_index(drop=True)
    plot_bar_scores(plot_df, treat_col, "WASPAS_Q", "WASPAS Combined Aggregation Ranking",
                     "WASPAS Q (\u03bb\u00b7WSM + (1-\u03bb)\u00b7WPM, higher = better)",
                     out_dir / "waspas_scores.png", color_palette="RdPu")
    plot_diagnostic_scatter(out_df_str, treat_col, "WSM_Q1", "WPM_Q2", "WASPAS_Q",
                             "Q\u2081 \u2014 Weighted Sum Model (additive, compensatory)",
                             "Q\u2082 \u2014 Weighted Product Model (multiplicative, non-compensatory)",
                             "WASPAS Q (combined, higher = better)",
                             "WASPAS: Additive vs. Multiplicative Aggregation",
                             "cividis", out_dir / "waspas_Q1Q2_scatter.png")
    plot_weighted_heatmap(R * weights_for_waspas, traits, data_df[treat_col].astype(str),
                           "Weighted Normalized Matrix (WASPAS)", out_dir / "waspas_heatmap.png")

    print(f"  -> WASPAS outputs saved to: {out_dir}")
    return {
        "method": "WASPAS", "treat_col": treat_col, "scores": Q,
        "treatments": data_df[treat_col].astype(str).values,
        "higher_is_better": True, "weights": weights_for_waspas,
    }


def run_edas_method(data_df, treat_col, traits, X, dir_list, weights_for_edas, out_dir):
    """Method 7 — EDAS distance-from-average ranking (same weights as Method 1)."""
    AS, SP, SN, PDA, NDA = run_edas_core(X, weights_for_edas, dir_list)

    out_df = data_df[[treat_col]].copy()
    out_df["SP"] = SP
    out_df["SN"] = SN
    out_df["EDAS_Score"] = AS
    out_df["rank"] = pd.Series(AS).rank(method="min", ascending=False).astype("Int64").values

    txt_path = out_dir / "scored_results_edas.txt"
    out_df.to_csv(txt_path, sep="\t", index=False, na_rep="NA", float_format="%.6f")
    summary = out_df.sort_values("EDAS_Score", ascending=False)
    with open(txt_path, "a", encoding="utf-8") as f:
        f.write("\n\n" + "=" * 55 + "\n")
        f.write("  RESULTS  (sorted by EDAS appraisal score -- HIGHER is better)\n")
        f.write("=" * 55 + "\n")
        f.write(summary[[treat_col, "SP", "SN", "EDAS_Score", "rank"]].to_string(index=False))
        f.write("\n")

    excel_path = out_dir / "scored_results_edas.xlsx"
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        out_df.to_excel(writer, sheet_name="Results", index=False)
        pd.DataFrame({"Trait": traits, "Direction": dir_list, "Weight_Used": weights_for_edas}).to_excel(
            writer, sheet_name="Weights", index=False)

    out_df_str = out_df.copy()
    out_df_str[treat_col] = out_df_str[treat_col].astype(str)

    plot_df = out_df_str[[treat_col, "EDAS_Score"]].sort_values("EDAS_Score", ascending=False).reset_index(drop=True)
    plot_bar_scores(plot_df, treat_col, "EDAS_Score", "EDAS Distance-from-Average Ranking",
                     "EDAS Appraisal Score (higher = better)", out_dir / "edas_scores.png",
                     color_palette="YlOrBr")
    plot_diagnostic_scatter(out_df_str, treat_col, "SP", "SN", "EDAS_Score",
                             "SP \u2014 positive distance from the average solution",
                             "SN \u2014 negative distance from the average solution",
                             "EDAS appraisal score (higher = better)",
                             "EDAS: Positive vs. Negative Distance from Average Solution",
                             "coolwarm_r", out_dir / "edas_SPSN_scatter.png")
    net_weighted = (PDA - NDA) * weights_for_edas
    plot_weighted_heatmap(net_weighted, traits, data_df[treat_col].astype(str),
                           "Net Weighted Distance-from-Average Matrix (EDAS)", out_dir / "edas_heatmap.png")

    print(f"  -> EDAS outputs saved to: {out_dir}")
    return {
        "method": "EDAS", "treat_col": treat_col, "scores": AS,
        "treatments": data_df[treat_col].astype(str).values,
        "higher_is_better": True, "weights": weights_for_edas,
    }


# ------------------------------------------------------------
#  Section 8 — Cross-method comparison
# ------------------------------------------------------------

def run_comparison(results: list, treat_col: str, out_dir: Path):
    main_dir = out_dir / "main_text"
    supp_dir = out_dir / "supplementary"
    main_dir.mkdir(parents=True, exist_ok=True)
    supp_dir.mkdir(parents=True, exist_ok=True)

    treatments = results[0]["treatments"]
    comp_df = pd.DataFrame({treat_col: treatments})

    goodness = {}
    ranks = {}
    for res in results:
        name = res["method"]
        scores = np.asarray(res["scores"], dtype=float)
        higher_is_better = res["higher_is_better"]

        comp_df[f"{name}_Score"] = scores
        rank_series = pd.Series(scores).rank(method="min", ascending=not higher_is_better).astype(int)
        comp_df[f"{name}_Rank"] = rank_series.values
        ranks[name] = rank_series.values

        s_min, s_max = np.min(scores), np.max(scores)
        rng = (s_max - s_min) if (s_max - s_min) != 0 else 1.0
        g = (scores - s_min) / rng if higher_is_better else 1.0 - (scores - s_min) / rng
        goodness[name] = g
        comp_df[f"{name}_Goodness"] = g

    method_names = [res["method"] for res in results]

    rank_matrix_df = pd.DataFrame(ranks)
    spearman_corr = rank_matrix_df.corr(method="spearman")

    rank_matrix = np.array([ranks[m] for m in method_names])
    W_overall = kendalls_w(rank_matrix)

    # ---- Axis-specific concordance (diagnostic: WHERE does disagreement come from) ----
    # Axis A = weighting scheme (same TOPSIS aggregation): Hybrid vs Simple vs CRITIC
    # Axis B = aggregation algorithm (same Hybrid AHP-Entropy weights): TOPSIS vs VIKOR
    #          vs MARCOS vs WASPAS vs EDAS
    weighting_axis = [m for m in ("Hybrid_AHP_Entropy", "Simple_TOPSIS", "CRITIC_AHP_TOPSIS") if m in method_names]
    aggregation_axis = [m for m in ("Hybrid_AHP_Entropy", "VIKOR", "MARCOS", "WASPAS", "EDAS") if m in method_names]

    W_weighting = (
        kendalls_w(np.array([ranks[m] for m in weighting_axis])) if len(weighting_axis) >= 2 else None
    )
    W_aggregation = (
        kendalls_w(np.array([ranks[m] for m in aggregation_axis])) if len(aggregation_axis) >= 2 else None
    )

    disc_rows = []
    for name in method_names:
        metrics = compute_discrimination_metrics(goodness[name])
        metrics["Method"] = name
        disc_rows.append(metrics)
    disc_df = pd.DataFrame(disc_rows)[["Method", "Range", "Std_Dev", "CV_percent", "Mean_Adjacent_Gap"]]
    # ---- NEW CODE: Consensus Ranking (Super-Ranking) ----
    # 1. محاسبه میانگین رتبه‌ها از بین تمام ۷ متد
    rank_cols = [f"{m}_Rank" for m in method_names]
    comp_df["Mean_Rank"] = comp_df[rank_cols].mean(axis=1)
    
    # 2. تعیین رتبه نهایی اجماعی (هرچه میانگین رتبه کمتر باشد، تیمار بهتر است)
    comp_df["Consensus_Rank"] = comp_df["Mean_Rank"].rank(method="min", ascending=True).astype(int)
    
    # 3. ایجاد یک دیتافریم مرتب‌شده بر اساس رتبه نهایی برای خروجی‌ها
    consensus_df = comp_df[[treat_col] + rank_cols + ["Mean_Rank", "Consensus_Rank"]].sort_values("Consensus_Rank")
    
    # 4. ذخیره رتبه‌بندی نهایی در یک فایل متنی مجزا
    with open(out_dir / "final_consensus_ranking.txt", "w", encoding="utf-8") as f:
        f.write("=" * 70 + "\n  FINAL CONSENSUS RANKING (SUPER-RANKING) ACROSS ALL 7 METHODS\n" + "=" * 70 + "\n\n")
        f.write(consensus_df.to_string(index=False) + "\n")

    # ---- Save data ----
    agreement_rows = [{"Metric": "Kendall's W -- ALL methods (overall concordance)", "Value": W_overall}]
    if W_weighting is not None:
        agreement_rows.append({
            "Metric": f"Kendall's W -- weighting axis ({', '.join(weighting_axis)})", "Value": W_weighting})
    if W_aggregation is not None:
        agreement_rows.append({
            "Metric": f"Kendall's W -- aggregation axis ({', '.join(aggregation_axis)})", "Value": W_aggregation})
    agreement_df = pd.DataFrame(agreement_rows)

    excel_path = out_dir / "comparison_data.xlsx"
    with pd.ExcelWriter(excel_path, engine="openpyxl") as writer:
        comp_df.to_excel(writer, sheet_name="Full_Comparison", index=False)
        spearman_corr.to_excel(writer, sheet_name="Spearman_Correlation")
        disc_df.to_excel(writer, sheet_name="Discrimination_Power", index=False)
        agreement_df.to_excel(writer, sheet_name="Overall_Agreement", index=False)
        consensus_df.to_excel(writer, sheet_name="Consensus_Ranking", index=False)

    with open(out_dir / "comparison_summary.txt", "w", encoding="utf-8") as f:
        f.write("=" * 60 + "\n  METHOD COMPARISON SUMMARY\n" + "=" * 60 + "\n\n")
        f.write(f"Methods compared ({len(method_names)}): {', '.join(method_names)}\n\n")
        f.write(f"Kendall's W -- ALL methods            = {W_overall:.4f}\n")
        if W_weighting is not None:
            f.write(f"Kendall's W -- weighting axis only     = {W_weighting:.4f}"
                    f"   ({', '.join(weighting_axis)})\n")
        if W_aggregation is not None:
            f.write(f"Kendall's W -- aggregation axis only   = {W_aggregation:.4f}"
                    f"   ({', '.join(aggregation_axis)})\n")
        f.write("  (0 = no agreement across methods, 1 = perfect agreement)\n")
        f.write("  Comparing the two axis-specific W values shows WHICH methodological\n")
        f.write("  choice -- how traits are weighted, or how scores are aggregated --\n")
        f.write("  contributes more to any disagreement in the final ranking.\n\n")
        f.write("Spearman Rank Correlation Matrix:\n")
        f.write(spearman_corr.round(3).to_string() + "\n\n")
        f.write("Discrimination Power (on normalized 0-1 goodness scores):\n")
        f.write(disc_df.round(4).to_string(index=False) + "\n\n")
        best_disc = disc_df.loc[disc_df["CV_percent"].idxmax(), "Method"]
        f.write(f"Method with the strongest treatment discrimination (highest CV%): {best_disc}\n")

    # ---- Plots ----
    plot_rank_bump_chart(comp_df, treat_col, method_names, main_dir / "rank_bump_chart.png")
    plot_goodness_grouped_bar(comp_df, treat_col, method_names, main_dir / "composite_scores_grouped_bar.png")
    plot_spearman_heatmap(spearman_corr, supp_dir / "spearman_correlation_heatmap.png")
    plot_discrimination_bar(disc_df, supp_dir / "discrimination_power_barplot.png")
    plot_rank_heatmap(comp_df, treat_col, method_names, supp_dir / "rank_heatmap.png")
    # ---- رسم نمودار رتبه‌بندی نهایی (Consensus) ----
    plot_df_consensus = consensus_df[[treat_col, "Mean_Rank"]].copy()
    plot_bar_scores(plot_df_consensus, treat_col, "Mean_Rank", 
                    "Final Consensus Ranking Across All 7 MCDM Methods",
                    "Mean Rank (Lower = Better)", 
                    main_dir / "final_consensus_ranking_bar.png",
                    color_palette="Blues", higher_is_better=False)
    

    print(f"  -> Comparison outputs saved to: {out_dir}")
    return comp_df, spearman_corr, disc_df, W_overall


# ------------------------------------------------------------
#  Section 9 — Main orchestration
# ------------------------------------------------------------

def main():
    base_out = Path(OUT_DIR)
    dir_hybrid = base_out / DIR_HYBRID
    dir_simple = base_out / DIR_SIMPLE
    dir_critic = base_out / DIR_CRITIC
    dir_vikor = base_out / DIR_VIKOR
    dir_marcos = base_out / DIR_MARCOS
    dir_waspas = base_out / DIR_WASPAS
    dir_edas = base_out / DIR_EDAS
    dir_compare = base_out / DIR_COMPARISON
    for d in (dir_hybrid, dir_simple, dir_critic, dir_vikor, dir_marcos, dir_waspas, dir_edas, dir_compare):
        d.mkdir(parents=True, exist_ok=True)

    # ── 1. Read Excel ──────────────────────────────────────
    print("Reading Excel file \u2026")
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
    traits = []
    ahp_w_raw = {}
    directions_raw = {}
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
        ahp_w_raw[name] = w
        directions_raw[name] = dir_char

    if not traits:
        print("[ERROR] No valid traits found. Check your Excel format.")
        sys.exit(1)
    print(f"\nTraits included in analysis ({len(traits)}): {traits}")

    # ── 3.5 Average replicates ─────────────────────────────
    print("\nProcessing data: converting to numeric and averaging replicates\u2026")
    data_df[treat_col] = data_df[treat_col].astype(str).str.strip()
    for t in traits:
        data_df[t] = pd.to_numeric(data_df[t], errors="coerce").fillna(0)
    original_len = len(data_df)
    data_df = data_df.groupby(treat_col, as_index=False)[traits].mean()
    print(f"Aggregated {original_len} rows into {len(data_df)} unique treatments.")

    # ── 4. Shared decision matrix (used by ALL 7 methods) ──
    X, directions_resolved = build_decision_matrix(data_df, traits, directions_raw, treat_col)
    dir_list = [directions_resolved[t] for t in traits]
    m = len(traits)

    # ── 5. Shared AHP weights ──────────────────────────────
    ahp_arr = np.array([ahp_w_raw[t] for t in traits], dtype=float)
    total_ahp = ahp_arr.sum()
    if total_ahp == 0:
        ahp_arr = np.ones(m)
        total_ahp = float(m)
    ahp_normalized = ahp_arr / total_ahp

    # ── 6. Shared Shannon Entropy weights (Method 1 only) ──
    shannon_arr = calculate_shannon_entropy_weights(X)

    apply_nature_plot_style()

    # ── 7. Run all seven methods ─────────────────────────────
    print("\n[1/7] Running Hybrid AHP\u2013Shannon Entropy TOPSIS (original method)\u2026")
    res_hybrid = run_hybrid_method(data_df, treat_col, traits, X, dir_list, ahp_normalized, shannon_arr, dir_hybrid)

    print("[2/7] Running Simple (equal-weight) TOPSIS\u2026")
    res_simple = run_simple_method(data_df, treat_col, traits, X, dir_list, dir_simple)

    print("[3/7] Running Hybrid AHP\u2013CRITIC TOPSIS\u2026")
    res_critic = run_critic_method(data_df, treat_col, traits, X, dir_list, ahp_normalized, dir_critic)

    print("[4/7] Running VIKOR (same weights as Method 1, different aggregation logic)\u2026")
    res_vikor = run_vikor_method(data_df, treat_col, traits, X, dir_list, res_hybrid["weights"], dir_vikor)

    print("[5/7] Running MARCOS (ideal/anti-ideal compromise utility)\u2026")
    res_marcos = run_marcos_method(data_df, treat_col, traits, X, dir_list, res_hybrid["weights"], dir_marcos)

    print("[6/7] Running WASPAS (additive-multiplicative aggregation)\u2026")
    res_waspas = run_waspas_method(data_df, treat_col, traits, X, dir_list, res_hybrid["weights"], dir_waspas)

    print("[7/7] Running EDAS (distance from average solution)\u2026")
    res_edas = run_edas_method(data_df, treat_col, traits, X, dir_list, res_hybrid["weights"], dir_edas)

    # ── 8. Cross-method comparison (now across all 7 methods) ──
    print("\nBuilding cross-method comparison\u2026")
    all_results = [res_hybrid, res_simple, res_critic, res_vikor, res_marcos, res_waspas, res_edas]
    comp_df, spearman_corr, disc_df, W = run_comparison(all_results, treat_col, dir_compare)

    # ── 9. Console summary ──────────────────────────────────
    print("\n" + "=" * 60)
    print("  ALL METHODS COMPLETE")
    print("=" * 60)
    print(f"Kendall's W (overall rank agreement across all 7 methods): {W:.4f}\n")
    rank_cols = [c for c in comp_df.columns if c.endswith("_Rank")]
    print(comp_df[[treat_col] + rank_cols].to_string(index=False))
    print(f"\nAll outputs saved under: {base_out}")
    print(f"  1) {dir_hybrid}")
    print(f"  2) {dir_simple}")
    print(f"  3) {dir_critic}")
    print(f"  4) {dir_vikor}")
    print(f"  5) {dir_marcos}")
    print(f"  6) {dir_waspas}")
    print(f"  7) {dir_edas}")
    print(f"  8) {dir_compare}  (main_text/ + supplementary/)")


if __name__ == "__main__":
    main()
