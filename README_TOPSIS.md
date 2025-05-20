
# TOPSIS - MCDM Web App

This is a Streamlit-based web app that implements the **TOPSIS (Technique for Order Preference by Similarity to Ideal Solution)** method for Multi-Criteria Decision Making (MCDM).

## 🚀 Features

- Upload standardized Excel files
- Handles benefit and cost criteria
- Computes normalized and weighted decision matrices
- Calculates ideal and anti-ideal solutions
- Ranks alternatives based on TOPSIS scores
- Allows downloading full results in Excel format

## 📥 Input Excel File Structure

Your Excel file must include **three sheets**:

### 1. `DecisionMatrix`
A table of alternatives (rows) and criteria (columns):

| Alternative   | Criteria 1 | Criteria 2 | Criteria 3 |
|---------------|------------|------------|------------|
| Alternative 1 | 7          | 8          | 6          |
| Alternative 2 | 9          | 7          | 8          |
| Alternative 3 | 6          | 5          | 7          |

### 2. `Weights`
A table of criteria and their corresponding weights:

| Criteria   | Weight |
|------------|--------|
| Criteria 1 | 0.4    |
| Criteria 2 | 0.3    |
| Criteria 3 | 0.3    |

### 3. `Types`
A table that defines the type of each criterion: `Benefit`, `Cost`, or `Target` (for future use):

| Criteria   | Type    |
|------------|---------|
| Criteria 1 | Benefit |
| Criteria 2 | Cost    |
| Criteria 3 | Benefit |

✅ **Note**: Weights must sum to 1.

## 🧾 Sample Files

- [Download Template Excel](TOPSIS_Template.xlsx)
- [Download App Script](app_topsis.py)
- [Example Table Image](TOPSIS_Table_Example.png)

## 📤 Output

- TOPSIS score and rank for each alternative
- Normalized matrix
- Weighted normalized matrix
- Ideal and Anti-Ideal solutions
- Downloadable Excel report

## 📦 How to Run the App

```bash
pip install streamlit pandas numpy xlsxwriter openpyxl
streamlit run app_topsis.py
```

---

Created by [Dr. Zahari Md Rodzi] for UiTM & Decision Making Research SIG.
