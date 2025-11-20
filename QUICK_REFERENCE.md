# Quick Reference Guide
## Excel File: 経営の診断書_フォーマット.xlsx
## Sheet: ４．３期比較表

---

## Table Layout Overview

```
┌─────────────────────────────────────────────┬─────────────────────────────────────────────┐
│  LEFT SECTION (Columns A-K)                 │  RIGHT SECTION (Columns M-W)                │
├─────────────────────────────────────────────┼─────────────────────────────────────────────┤
│                                             │                                             │
│  ① 売上高内訳表                              │  ② 販売費及び一般管理費比較表                  │
│     Rows 3-9                                │     Rows 3-30                               │
│     Data: 6-8 (3 rows)                      │     Data: 6-11 (6 rows)                     │
│                                             │     Total: Row 12                           │
│                                             │                                             │
├─────────────────────────────────────────────┤                                             │
│                                             │                                             │
│  ③ 変動費内訳比較表                           │                                             │
│     Rows 10-20                              │                                             │
│     Data: 14-19 (6 rows)                    │                                             │
│     Subtotal: Row 17                        │                                             │
│     Total: Row 20                           │                                             │
│                                             │                                             │
├─────────────────────────────────────────────┼─────────────────────────────────────────────┤
│                                             │                                             │
│  ④ 製造経費比較表                             │  ⑤ その他損益比較表                           │
│     Rows 21-41                              │     Rows 31-44                              │
│     Data: 25-39 (15 rows)                   │     Data: 35-44 (10 rows)                   │
│     Subtotal 1: Row 30                      │     Total: Row 44                           │
│     Subtotal 2: Row 40                      │                                             │
│     Total: Row 41                           │                                             │
│                                             │                                             │
└─────────────────────────────────────────────┴─────────────────────────────────────────────┘
```

---

## Column Structure - Standard Layout (A-K)

Used by: **Tables ①, ③, ④**

```
┌───┬─────────┬─────────┬─────────┬─────────┬─────────┬─────────┬─────────┬─────────┬─────────┬─────────┐
│ A │    B    │    C    │    D    │    E    │    F    │    G    │    H    │    I    │    J    │    K    │
├───┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┤
│項目│  前々期  │  前々期  │   前期   │   前期   │   前期   │   前期   │   当期   │   当期   │   当期   │   当期   │
│名 │   金額   │  売上比  │   金額   │  売上比  │  前年比  │ 対年変更 │   金額   │  売上比  │  前年比  │ 対年変更 │
└───┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┘
```

---

## Column Structure - Extended Layout (M-W)

Used by: **Tables ②, ⑤**

```
┌───┬─────────┬─────────┬─────────┬─────────┬─────────┬─────────┬─────────┬─────────┬─────────┬─────────┐
│ M │    N    │    O    │    P    │    Q    │    R    │    S    │    T    │    U    │    V    │    W    │
├───┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┼─────────┤
│項目│  前々期  │  前々期  │   前期   │   前期   │   前期   │   前期   │   当期   │   当期   │   当期   │   当期   │
│名 │   金額   │  売上比  │   金額   │  売上比  │  前年比  │ 対年変更 │   金額   │  売上比  │  前年比  │ 対年変更 │
└───┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┴─────────┘
```

---

## Table Details

### ① 売上高内訳表 (Sales Breakdown)
- **Location:** A3:K9
- **Header Row:** 5
- **Data Rows:** 6-8 (3 rows)
- **Total Row:** 9
- **Items:**
  - Row 6: 商品売上
  - Row 7-8: (Available for data)
  - Row 9: 合計

### ② 販売費及び一般管理費比較表 (SG&A)
- **Location:** M3:W30+
- **Header Row:** 5
- **Data Rows:** 6-11 (6 rows for personnel costs)
- **Subtotal Row:** 12 (人件費小計)
- **Items:**
  - Row 6: 役員報酬 (Executive Compensation)
  - Row 7: 給与手当 (Salaries)
  - Row 8: 賞与 (Bonuses)
  - Row 9: 法定福利費 (Statutory Welfare)
  - Row 10: 福利厚生費 (Welfare Expenses)
  - Row 11: 退職金 (Retirement Benefits)
  - Row 12: （人件費小計） - Subtotal
  - Row 13+: Additional SG&A items

### ③ 変動費内訳比較表 (Variable Costs)
- **Location:** A10:K20
- **Header Row:** 13
- **Data Rows:** 14-19 (6 rows)
- **Subtotal Row:** 17 (小計)
- **Total Row:** 20
- **Items:**
  - Row 14: 仕入高 (Purchases)
  - Row 15: 外注費 (Outsourcing)
  - Row 16: (Available for data)
  - Row 17: （小計） - Subtotal
  - Row 18: 期首棚卸高 (Beginning Inventory)
  - Row 19: 期末棚卸高 (Ending Inventory)
  - Row 20: 合計 - Total

### ④ 製造経費比較表 (Manufacturing Overhead)
- **Location:** A21:K41
- **Header Row:** 24
- **Data Rows:** 25-39 (15 rows)
- **Subtotal Row 1:** 30 (労務費計)
- **Subtotal Row 2:** 40 (経費計)
- **Total Row:** 41
- **Sections:**
  - **Labor Costs (25-30):**
    - Row 25: 賃金給料 (Wages)
    - Row 26: 賞与 (Bonuses)
    - Row 27: 法定福利費 (Statutory Welfare)
    - Row 28: 福利厚生費 (Welfare)
    - Row 30: （労務費計） - Labor Subtotal
  - **Expenses (31-40):**
    - Row 31: 減価償却費 (Depreciation)
    - Row 32-38: Other expenses
    - Row 39: その他経費 (Other Expenses)
    - Row 40: （経費計） - Expense Subtotal
  - Row 41: 合計 - Total

### ⑤ その他損益比較表 (Other P&L)
- **Location:** M31:W44
- **Header Row:** 34
- **Data Rows:** 35-44 (10 rows)
- **Total Row:** 44
- **Items:**
  - Row 35: 売上総利益 (Gross Profit)
  - Row 36: 営業利益 (Operating Profit)
  - Row 37: 営業外収益 (Non-Operating Income)
  - Row 38: 営業外費用 (Non-Operating Expenses)
  - Row 39: 経常利益 (Ordinary Profit)
  - Row 40: 特別利益 (Extraordinary Income)
  - Row 41: 特別損失 (Extraordinary Losses)
  - Row 42: 税引前利益 (Profit Before Tax)
  - Row 43: 法人税等 (Corporate Tax)
  - Row 44: 当期純利益 (Net Profit)

---

## Programming Reference

### Reading Data Example (Python with openpyxl)

```python
import openpyxl

wb = openpyxl.load_workbook('経営の診断書_フォーマット.xlsx')
ws = wb['４．３期比較表']

# Table ① - 売上高内訳表
for row in range(6, 9):  # Rows 6-8
    item_name = ws[f'A{row}'].value
    前々期_金額 = ws[f'B{row}'].value
    前々期_売上比 = ws[f'C{row}'].value
    前期_金額 = ws[f'D{row}'].value
    前期_売上比 = ws[f'E{row}'].value
    前期_前年比 = ws[f'F{row}'].value
    前期_対年変更 = ws[f'G{row}'].value
    当期_金額 = ws[f'H{row}'].value
    当期_売上比 = ws[f'I{row}'].value
    当期_前年比 = ws[f'J{row}'].value
    当期_対年変更 = ws[f'K{row}'].value

# Table ② - 販売費及び一般管理費比較表
for row in range(6, 12):  # Rows 6-11
    item_name = ws[f'M{row}'].value
    前々期_金額 = ws[f'N{row}'].value
    前々期_売上比 = ws[f'O{row}'].value
    前期_金額 = ws[f'P{row}'].value
    前期_売上比 = ws[f'Q{row}'].value
    前期_前年比 = ws[f'R{row}'].value
    前期_対年変更 = ws[f'S{row}'].value
    当期_金額 = ws[f'T{row}'].value
    当期_売上比 = ws[f'U{row}'].value
    当期_前年比 = ws[f'V{row}'].value
    当期_対年変更 = ws[f'W{row}'].value
```

### Writing Data Example

```python
# Write to Table ① - Row 7 (empty row available for data)
row = 7
ws[f'A{row}'] = "新商品売上"  # Item name
ws[f'B{row}'] = 1000000  # 前々期 金額
ws[f'D{row}'] = 1200000  # 前期 金額
ws[f'H{row}'] = 1500000  # 当期 金額

# Note: 売上比, 前年比, 対年変更 columns contain formulas
# These will auto-calculate when the file is opened in Excel

wb.save('経営の診断書_フォーマット_updated.xlsx')
```

---

## Important Notes

1. **Formulas:** Columns for 売上比, 前年比, and 対年変更 contain formulas. Only write to 金額 columns.

2. **Reference Cell:** Most percentage formulas reference cell B9 (Table ① total) as the base.

3. **Empty Rows:** Several rows are empty and available for additional data entries.

4. **Merged Cells:** Table titles use merged cells. Be careful when modifying these areas.

5. **Unit:** All amounts are in 千円 (thousands of yen).

---

## File Paths

- **Excel File:** `c:\Users\t.nohora\OneDrive - 御堂筋税理士法人\デスクトップ\MyGasProject\経営の診断書_フォーマット.xlsx`
- **JSON Structure:** `c:\Users\t.nohora\OneDrive - 御堂筋税理士法人\デスクトップ\MyGasProject\excel_structure_complete.json`
- **Full Report:** `c:\Users\t.nohora\OneDrive - 御堂筋税理士法人\デスクトップ\MyGasProject\EXCEL_STRUCTURE_REPORT.md`
- **This Guide:** `c:\Users\t.nohora\OneDrive - 御堂筋税理士法人\デスクトップ\MyGasProject\QUICK_REFERENCE.md`
