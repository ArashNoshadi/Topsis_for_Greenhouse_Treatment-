
# Hybrid TOPSIS Analysis for Ranking Plant-Parasitic Nematode Treatments

This project provides a robust Python implementation of **TOPSIS** (Technique for Order of Preference by Similarity to Ideal Solution), enhanced with a **Hybrid Weighting System**.

Designed for agricultural and biological research, this tool ranks the efficacy of treatments (e.g., organic amendments, bioproducts, chemical controls) against plant-parasitic nematodes (like *Meloidogyne javanica*). It goes beyond simple ranking by integrating subjective expert opinions, objective data variance, and model reliability into a single comprehensive analysis.

## 🚀 Key Features

### 🧠 Advanced Hybrid Weighting Engine
Unlike standard TOPSIS calculators, this tool computes the final weight for each trait by combining three distinct factors:
1.  **Subjective Weights (AHP-based)**: User-defined weights reflecting expert priorities (read from Excel).
2.  **Objective Weights (Shannon Entropy)**: Automatically calculated based on the data's information content (divergence).
3.  **Reliability Factor (Error Correction)**: Penalizes traits with high predictive errors (using an inverse function: $W \propto \frac{1}{1 + Error}$).

### 📊 Full TOPSIS Implementation
- Calculates Normalized Decision Matrix, Weighted Normalized Matrix, PIS (Positive Ideal Solution), and NIS (Negative Ideal Solution).
- Generates a final **Composite Score ($C_i$)** for ranking.

### 🛠️ Automated & Configurable
- **Excel Integration**: Reads data, weights, directions, and error metrics directly from a single `.xlsx` file.
- **Smart Direction Handling**:
    - **Benefit (+)**: Maximizes traits like 'SFW', 'SDW', 'SL'.
    - **Cost (-)**: Minimizes traits like 'Eggs', 'Galls', 'RF'.
- **Case-Insensitive**: Automatically detects columns regardless of capitalization (e.g., "treatment", "Treatment").

### 📈 Publication-Ready Outputs
- **Detailed Excel Report**: Includes a dedicated **'Weights'** sheet showing the breakdown of AHP, Entropy, and Reliability weights for every trait.
- **Visuals**: Generates a high-resolution (600 DPI) bar chart of the rankings.

## 🛠️ Requirements

The script requires the following Python libraries:

- `pandas`
- `numpy`
- `matplotlib`
- `seaborn`
- `openpyxl`

### Installation

```bash
pip install pandas numpy matplotlib seaborn openpyxl
````

## ⚙️ How to Use

### 1\. Prepare the Input Excel File

Your `Input.xlsx` file must contain a sheet with the following structure:

  - **Treatment**: Names of the treatments (e.g., 'Control', 'Tervigo').
  - **Value Columns**: Raw data for each trait (e.g., `RF`, `Eggs`, `Gall`).
  - **Weight Columns** (Optional): AHP/Subjective weights (e.g., `RF-wgt`).
  - **Direction Columns**:
      - `+` for benefit (higher is better).
      - `-` for cost (lower is better).
  - **Error Columns** (Optional): Model error metrics (e.g., `RF-err`). *Traits with higher errors will receive lower influence weights.*

### 2\. Configure the Script

Open the Python script and update the **Configuration** block at the top:

```python
# File Paths
IN_PATH = "data/Input.xlsx"   # Path to your input file
OUT_DIR = "results/v1"        # Output folder

# Trait Mapping
# Map your traits to their corresponding columns in Excel
TRAITS_RAW = {
    "RF":   {"weight_col": "RF-wgt", "direction_col": "RF-drc", "error_col": "RF-err"},
    "Eggs": {"weight_col": "Eggs-wgt", "direction_col": "Eggs-drc", "error_col": "Eggs-err"},
    # ... add other traits here
}
```

### 3\. Run the Analysis

```bash
python hybrid_topsis.py
```

## 📊 Outputs

Upon execution, the script generates three files in your output directory:

1.  **`scored_results.xlsx`**:
      - **Sheet 'Results'**: The main ranking table with composite scores, ranks, and normalized values.
      - **Sheet 'Weights'**: A transparent breakdown of how the final weights were derived (AHP × Shannon × Reliability).
2.  **`scored_results.txt`**: A tab-separated text file of the main results for easy import into other statistical software (SAS, SPSS, R).
3.  **`composite_scores.png`**: A professional horizontal bar chart visualizing the treatment rankings, sorted from best to worst.

## 📝 Methodology Brief

The final weight ($W_j$) for each trait is calculated as:

$$ W_{final_j} = \frac{W_{AHP_j} \times W_{Shannon_j} \times \frac{1}{1 + Error_j}}{\sum (W_{AHP} \times W_{Shannon} \times \frac{1}{1 + Error})} $$

This ensures that the final ranking respects expert opinion, data structure, and model reliability simultaneously.

## 📫 Author & Contact

  - **Developed by**: Arash Noshadi
  - **Contact**: Nowshadiarash@gmail.com
  - **License**: MIT License
  - **DOI**: *[Insert DOI here]*

