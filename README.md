# 💰 Cash Report - Target Wise Daily Achievement

This project generates a **Cash Deposit Target Report** by comparing daily achieved amounts with predefined targets for each ASM.

It takes:
- An **Export Excel file** (with sheet `Table`)  
- A **Mapping Excel file** (with sheet `MAPPING`)  

And produces:
- A styled Excel report (`Cash_Report.xlsx`) showing **Target, Amount, and % Achievement** (with Total row).  

---

## ⚡ Features
- Reads export data (`deposit_date`, `type`, `status`, `amount`, `code`)  
- Filters only latest day records (type = *cash/card*, status = *accepted*)  
- Merges with target mapping (ASM wise)  
- Calculates:
  - Total **Target**
  - Total **Achieved Amount**
  - **% Achievement**
- Adds a **TOTAL** row
- Saves a styled Excel file with:
  - Chocolate background headers
  - White bold fonts
  - Borders for all cells
  - Auto column width
  - Centered report title

---

## 📦 Requirements
- Python 3.x  
- Libraries:
  ```bash
  pip install pandas openpyxl
