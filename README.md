# TOPSIS Analysis for Ranking Plant-Parasitic Nematode Treatments

This project provides a Python script for TOPSIS (Technique for Order of Preference by Similarity to Ideal Solution), a multi-criteria decision-making (MCDM) method.

This script is specifically configured to analyze experimental data for ranking the efficacy of various treatments (such as organic amendments, bioproducts, or chemical controls) against plant-parasitic nematodes (like Meloidogyne javanica). It evaluates treatments based on multiple criteria—such as nematode reproduction indices and plant growth parameters—to generate a single, comprehensive "composite score" for ranking.

## 🚀 Features

- **Full TOPSIS Implementation**: Calculates the normalized decision matrix, weighted normalized matrix, Positive Ideal Solution (PIS), Negative Ideal Solution (NIS), and the final relative closeness (Ci) score.
- **Reads from Excel**: Directly ingests data from a specified `Input.xlsx` file and sheet.
- **Easy Configuration**: All inputs, outputs, and trait definitions are centralized in a configuration block at the top of the script.
- **Automatic Weight & Direction**: Reads the weight (importance) and direction (benefit/cost) for each criterion directly from the input Excel file.
  - **Benefit (+)**: Criteria to be maximized (e.g., 'SFW', 'SDW', 'SL').
  - **Cost (-)**: Criteria to be minimized (e.g., 'Eggs', 'Galls', 'Egg Masses', 'RF').
- **Equal Weighting Fallback**: If no weight columns are provided, the script automatically assigns equal weights to all criteria.
- **Comprehensive Outputs**:
  - **.xlsx file**: A multi-sheet Excel file with the full ranked results and a separate sheet detailing the final weights and directions used in the analysis.
  - **.txt file**: A tab-separated text file of the primary results for easy import into other software.
  - **.png plot**: A publication-quality horizontal bar chart visualizing the composite scores, sorted from best to worst.
- **Case-Insensitive**: Automatically finds column names (e.g., "Treatment") regardless of capitalization.

## 🛠️ Requirements

The script requires the following Python libraries. You can install them all using the `requirements.txt` file (if provided) or individually.

- pandas
- numpy
- matplotlib
- seaborn
- openpyxl (for reading/writing .xlsx files)

### Installation

```bash
pip install pandas numpy matplotlib seaborn openpyxl
```

## ⚙️ How to Use

1. **Install Requirements**:

   ```bash
   pip install pandas numpy matplotlib seaborn openpyxl
   ```

2. **Prepare the Input.xlsx file**: Your data file must be structured correctly.
   - **Treatment Column**: A column listing the names of each treatment (e.g., 'Untreated Control', 'Tervigo', 'Fish Bone Meal 4%', 'B. subtilis').
   - **Value Columns**: Columns for the data of each criterion (e.g., RF, Eggs, Gall, SDW, SFW). The names must match the keys in the `TRAITS_RAW` dictionary.
   - **Weight Columns**: (Optional) Columns for the weight of each criterion (e.g., RF-wgt, Eggs-wgt). This is typically a constant number down the entire column.
   - **Direction Columns**: Columns for the goal of each criterion (e.g., RF-drc, Eggs-drc). This column should contain either:
     - `+` for benefit criteria (higher is better, e.g., 'SFW' - Shoot Fresh Weight).
     - `-` for cost criteria (lower is better, e.g., 'Galls' or 'RF' - Reproduction Factor).

3. **Configure the Script**: Open the Python (.py) file and edit the Configuration section at the top.
   - `IN_PATH`: The file path to your `Input.xlsx` file.
   - `SHEET_NAME`: The name of the sheet in the Excel file that contains your data.
   - `OUT_DIR`: The path to the folder where all output files will be saved.
   - `TRAITS_RAW`: This is the most important part. Define all your criteria here, mapping the trait name to its corresponding weight and direction columns in the Excel file.

   ⚠️ **Important**: The script uses absolute paths (e.g., `G:\...`). This will not work on any other computer. It is highly recommended to use relative paths.

   **Example (Recommended)**: Create a `data` folder for your input and an `output` folder for your results.

   ```python
   IN_PATH = "data/Input.xlsx"
   OUT_DIR = "output/v1_results"
   ```

4. **Run the Script**:

   ```bash
   python your_script_name.py
   ```

## 📊 Outputs

After running, the following files will be created in your specified `OUT_DIR`:

- **scored_results.xlsx**:
  - **Sheet 'Results'**: Contains the Treatment, composite_score, rank, and all normalized/weighted values for each criterion.
  - **Sheet 'Weights'**: A summary of the final weights and directions applied to each trait.
- **scored_results.txt**: A tab-delimited text version of the 'Results' sheet, ideal for quick review or data import.
- **composite_scores.png**: A high-DPI (600 dpi) horizontal bar plot showing the final composite_score for each treatment, sorted in descending order (best treatment at the top).

## 📫 Author & Contact
- **Developed by**: Arash Noshadi

- **Contact**: Nowshadiarash@gmail.com

*License: MIT License*

*Doi*: ****
