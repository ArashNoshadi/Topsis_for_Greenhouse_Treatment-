📘 README — TOPSIS Analysis for Ranking Plant-Parasitic Nematode Treatments
📖 Overview
This project implements a full TOPSIS (Technique for Order of Preference by Similarity to Ideal Solution) analysis for ranking experimental treatments used against plant-parasitic nematodes such as Meloidogyne javanica.

It provides a transparent, reproducible, and flexible framework for determining which treatments (organic amendments, bioproducts, or chemical controls) perform best based on multiple biological and plant growth criteria.

🚀 Features
Full TOPSIS Implementation:

Normalization of the decision matrix
Weighted normalization
Computation of Positive Ideal Solution (PIS) and Negative Ideal Solution (NIS)
Calculation of the final composite score (Ci) and ranking
Excel-based Input:

Reads raw data, weights, and directions directly from a single Excel file.
Configurable:

Central configuration block for all paths, sheet names, and criteria mappings.
Automatic Weight & Direction Detection:

Reads weight and benefit/cost direction from the Excel file.
Defaults to equal weights when not specified.
Comprehensive Outputs:

.xlsx: Detailed results and weights
.txt: Tab-delimited summary
.png: High-quality bar plot of ranked treatments
Case-Insensitive Column Detection:

Automatically recognizes treatment and criterion columns regardless of capitalization.
🧩 Criteria Types
Benefit (“+”) Criteria: Higher values are better
Examples: SFW, SDW, SL
Cost (“–”) Criteria: Lower values are better
Examples: Eggs, Galls, Egg Masses, RF
🛠️ Requirements
Install dependencies using:


content_copy
text
pip install pandas numpy matplotlib seaborn openpyxl
Required Python Packages:

pandas
numpy
matplotlib
seaborn
openpyxl
⚙️ Configuration
Modify these variables at the top of the script before running:


content_copy
python

note_add
ویرایش با Canvas
# Example Configuration
IN_PATH = "data/Input.xlsx"
SHEET_NAME = "Sheet1"
OUT_DIR = "output/v1_results"

# Define trait names, their weights, and directions
TRAITS_RAW = {
    "RF": {"weight_col": "RF-wgt", "direction_col": "RF-drc"},
    "Eggs": {"weight_col": "Eggs-wgt", "direction_col": "Eggs-drc"},
    "Galls": {"weight_col": "Galls-wgt", "direction_col": "Galls-drc"},
    "SFW": {"weight_col": "SFW-wgt", "direction_col": "SFW-drc"},
    "SDW": {"weight_col": "SDW-wgt", "direction_col": "SDW-drc"},
}
📌 Important:

Use relative paths (not absolute paths) to keep the project portable.

Recommended project structure:


content_copy
text
├── data/
│   └── Input.xlsx
├── output/
│   ├── scored_results.xlsx
│   ├── scored_results.txt
│   └── composite_scores.png
└── TOPSIS_analysis.py
📂 Input File Structure
Your Excel input (Input.xlsx) should contain:

| Treatment | RF | Eggs | Galls | SDW | SFW | RF-wgt | RF-drc | … |

|------------|----|-------|--------|------|------|----------|---------|

| Control | 10.5 | 5200 | 75 | 1.2 | 3.0 | 1 | - | … |

| Tervigo | 6.0 | 2000 | 20 | 2.1 | 4.0 | 1 | - | … |

| Fish Bone 4% | 8.0 | 3000 | 45 | 1.8 | 3.8 | 1 | - | … |

▶️ How to Run
Install dependencies

content_copy
text
   pip install pandas numpy matplotlib seaborn openpyxl
Run the analysis

content_copy
text
   python TOPSIS_analysis.py
Check outputs in the folder specified by OUT_DIR.
📊 Output Files
File	Description
scored_results.xlsx	Full ranked results (sheet: Results) and applied weights/directions (sheet: Weights)
scored_results.txt	Main results in tab-delimited text format
composite_scores.png	Sorted horizontal bar chart of final composite (Ci) scores
🧮 Method Summary (TOPSIS Steps)
Normalization of decision matrix
Weighted normalization using provided or equal weights
Determination of PIS and NIS
Distance calculation from ideal and anti-ideal points
Relative closeness (Ci) computation:
𝐶
𝑖
=
𝐷
𝑖
−
𝐷
𝑖
+
+
𝐷
𝑖
−
C 
i
​
 = 
D 
i
+
​
 +D 
i
−
​
 
D 
i
−
​
 
​
 
Ranking based on descending C_i values (higher = better treatment)
📈 Visualization
The script automatically generates a publication-quality horizontal bar chart (composite_scores.png) showing the treatments ranked from best to worst based on their composite scores.

🧾 Citation
If you use this script or its methodology in academic work, please cite it as:

TOPSIS Analysis for Ranking Plant-Parasitic Nematode Treatments (2025).

Python implementation and documentation by [Your Name or Organization].

📫 Author & Contact
Developed by: Arash Noshadi

Contact: Nowshadiarash@gmail.com

License: MIT License
