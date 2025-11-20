# Excel Structure Analysis Report
## File: 経営の診断書_フォーマット.xlsx
## Sheet: ４．３期比較表 (3-Period Comparison Table)

---

## Overview

**Sheet Name:** ４．３期比較表
**Unit:** 千円 (thousands of yen)
**Number of Tables:** 5 financial comparison tables

---

## Column Structure Reference

### Standard Column Layout (Tables ①, ③, ④)
Located in **Columns A-K**

| Column | Period | Data Type | Description |
|--------|--------|-----------|-------------|
| A | - | Item Name | 項目名 |
| B | 前々期 | 金額 | Amount (Two Periods Ago) |
| C | 前々期 | 売上比 | Sales Ratio (%) |
| D | 前期 | 金額 | Amount (Previous Period) |
| E | 前期 | 売上比 | Sales Ratio (%) |
| F | 前期 | 前年比 | Year-over-Year Ratio (%) |
| G | 前期 | 対年変更 | Year-over-Year Change |
| H | 当期 | 金額 | Amount (Current Period) |
| I | 当期 | 売上比 | Sales Ratio (%) |
| J | 当期 | 前年比 | Year-over-Year Ratio (%) |
| K | 当期 | 対年変更 | Year-over-Year Change |

### Extended Column Layout (Tables ②, ⑤)
Located in **Columns M-W**

| Column | Period | Data Type | Description |
|--------|--------|-----------|-------------|
| M | - | Item Name | 項目名 |
| N | 前々期 | 金額 | Amount (Two Periods Ago) |
| O | 前々期 | 売上比 | Sales Ratio (%) |
| P | 前期 | 金額 | Amount (Previous Period) |
| Q | 前期 | 売上比 | Sales Ratio (%) |
| R | 前期 | 前年比 | Year-over-Year Ratio (%) |
| S | 前期 | 対年変更 | Year-over-Year Change |
| T | 当期 | 金額 | Amount (Current Period) |
| U | 当期 | 売上比 | Sales Ratio (%) |
| V | 当期 | 前年比 | Year-over-Year Ratio (%) |
| W | 当期 | 対年変更 | Year-over-Year Change |

---

## Table 1: ① 売上高内訳表 (Sales Breakdown Table)

### Location
- **Section:** Left (Columns A-K)
- **Header Row:** 5
- **Data Range:** Rows 6-8
- **Total Row:** 9

### Structure Details
```
Row 3: Table title
Row 4: Period headers (前々期, 前期, 当期)
Row 5: Column headers (金額, 売上比, 前年比, 対年変更)
Rows 6-8: Data rows
Row 9: Total (合計)
```

### Data Items
| Row | Item Name | Type |
|-----|-----------|------|
| 6 | 商品売上 | Data |
| 7 | (Empty - available for data) | Data |
| 8 | (Empty - available for data) | Data |
| 9 | 合計 | Total |

### Available Data Rows: 3

### Column Mapping
- **Item Names:** Column A
- **前々期:** Columns B (金額), C (売上比)
- **前期:** Columns D (金額), E (売上比), F (前年比), G (対年変更)
- **当期:** Columns H (金額), I (売上比), J (前年比), K (対年変更)

---

## Table 2: ② 販売費及び一般管理費比較表 (SG&A Comparison Table)

### Location
- **Section:** Right (Columns M-W)
- **Header Row:** 5
- **Data Range:** Rows 6-11
- **Total Row:** 12

### Structure Details
```
Row 3: Table title (merged M3:P3)
Row 4: Period headers (前々期, 前期, 当期)
Row 5: Column headers (金額, 売上比, 前年比, 対年変更)
Rows 6-11: Data rows
Row 12: Subtotal (人件費小計)
```

### Data Items
| Row | Item Name | Type |
|-----|-----------|------|
| 6 | 役員報酬 | Data |
| 7 | 給与手当 | Data |
| 8 | 賞与 | Data |
| 9 | 法定福利費 | Data |
| 10 | 福利厚生費 | Data |
| 11 | 退職金 | Data |
| 12 | （人件費小計） | Subtotal |

### Additional Data Rows
- Row 13: 減価償却費
- Row 14-23: Other SG&A items
- Row 24-28: More items
- Row 29: その他
- Row 30: 合計

### Available Data Rows: 6 (primary items) + additional rows

### Column Mapping
- **Item Names:** Column M
- **前々期:** Columns N (金額), O (売上比)
- **前期:** Columns P (金額), Q (売上比), R (前年比), S (対年変更)
- **当期:** Columns T (金額), U (売上比), V (前年比), W (対年変更)

---

## Table 3: ③ 変動費内訳比較表 (Variable Cost Breakdown Table)

### Location
- **Section:** Left (Columns A-K)
- **Header Row:** 13
- **Data Range:** Rows 14-19
- **Subtotal Row:** 17
- **Total Row:** 20

### Structure Details
```
Row 10-11: Table title (merged A10:D11)
Row 12: Period headers (前々期, 前期, 当期)
Row 13: Column headers (金額, 売上比, 前年比, 対年変更)
Rows 14-16: Primary data rows
Row 17: Subtotal (小計)
Rows 18-19: Inventory adjustment rows
Row 20: Total (合計)
```

### Data Items
| Row | Item Name | Type |
|-----|-----------|------|
| 14 | 仕入高 | Data |
| 15 | 外注費 | Data |
| 16 | (Empty - available for data) | Data |
| 17 | （小計） | Subtotal |
| 18 | 期首棚卸高 | Adjustment |
| 19 | 期末棚卸高 | Adjustment |
| 20 | 合計 | Total |

### Available Data Rows: 3 primary + 2 adjustment rows

### Column Mapping
- **Item Names:** Column A
- **前々期:** Columns B (金額), C (売上比)
- **前期:** Columns D (金額), E (売上比), F (前年比), G (対年変更)
- **当期:** Columns H (金額), I (売上比), J (前年比), K (対年変更)

---

## Table 4: ④ 製造経費比較表 (Manufacturing Overhead Table)

### Location
- **Section:** Left (Columns A-K)
- **Header Row:** 24
- **Data Range:** Rows 25-39
- **Subtotal Row 1:** 30 (労務費計)
- **Subtotal Row 2:** 40 (経費計)
- **Total Row:** 41

### Structure Details
```
Row 21-22: Table title (merged A21:D22)
Row 23: Period headers (前々期, 前期, 当期)
Row 24: Column headers (金額, 売上比, 前年比, 対年変更)
Rows 25-29: Labor cost items
Row 30: Labor cost subtotal (労務費計)
Rows 31-39: Other manufacturing overhead items
Row 40: Expense subtotal (経費計)
Row 41: Total (合計)
```

### Data Items
| Row | Item Name | Type |
|-----|-----------|------|
| 25 | 賃金給料 | Data (Labor) |
| 26 | 賞与 | Data (Labor) |
| 27 | 法定福利費 | Data (Labor) |
| 28 | 福利厚生費 | Data (Labor) |
| 29 | (Empty row) | Data (Labor) |
| 30 | （労務費計） | Subtotal |
| 31 | 減価償却費 | Data (Expense) |
| 32-38 | (Various expense items) | Data (Expense) |
| 39 | その他経費 | Data (Expense) |
| 40 | （経費計） | Subtotal |
| 41 | 合計 | Total |

### Available Data Rows: 15 (5 labor + 10 expense items)

### Column Mapping
- **Item Names:** Column A
- **前々期:** Columns B (金額), C (売上比)
- **前期:** Columns D (金額), E (売上比), F (前年比), G (対年変更)
- **当期:** Columns H (金額), I (売上比), J (前年比), K (対年変更)

---

## Table 5: ⑤ その他損益比較表 (Other P&L Comparison Table)

### Location
- **Section:** Right (Columns M-W)
- **Header Row:** 34
- **Data Range:** Rows 35-44
- **Total Row:** 44

### Structure Details
```
Row 31-32: Table title (merged M31:P32)
Row 33: Period headers (前々期, 前期, 当期)
Row 34: Column headers (金額, 売上比, 前年比, 対年変更)
Rows 35-44: Data rows (P&L summary items)
Row 44: Final total (当期純利益)
```

### Data Items
| Row | Item Name | Type |
|-----|-----------|------|
| 35 | 売上総利益 | Data |
| 36 | 営業利益 | Data |
| 37 | 営業外収益 | Data |
| 38 | 営業外費用 | Data |
| 39 | 経常利益 | Data |
| 40 | 特別利益 | Data |
| 41 | 特別損失 | Data |
| 42 | 税引前利益 | Data |
| 43 | 法人税等 | Data |
| 44 | 当期純利益 | Total |

### Available Data Rows: 10

### Column Mapping
- **Item Names:** Column M
- **前々期:** Columns N (金額), O (売上比)
- **前期:** Columns P (金額), Q (売上比), R (前年比), S (対年変更)
- **当期:** Columns T (金額), U (売上比), V (前年比), W (対年変更)

---

## Summary Table

| Table | Name | Location | Header Row | Data Rows | Total/Subtotal Rows | Row Count |
|-------|------|----------|------------|-----------|---------------------|-----------|
| ① | 売上高内訳表 | A-K | 5 | 6-8 | 9 | 3 |
| ② | 販売費及び一般管理費比較表 | M-W | 5 | 6-11 | 12 | 6+ |
| ③ | 変動費内訳比較表 | A-K | 13 | 14-19 | 17, 20 | 6 |
| ④ | 製造経費比較表 | A-K | 24 | 25-39 | 30, 40, 41 | 15 |
| ⑤ | その他損益比較表 | M-W | 34 | 35-44 | 44 | 10 |

---

## Key Findings

1. **Sheet Structure:** The sheet contains 5 financial comparison tables organized in two vertical sections
   - Left section (Columns A-K): Tables ①, ③, ④
   - Right section (Columns M-W): Tables ②, ⑤

2. **Column Pattern:** All tables follow the same column structure with 11 columns:
   - 1 item name column
   - 10 data columns (金額 and 売上比 for each of 3 periods, plus 前年比 and 対年変更 for 前期 and 当期)

3. **Period Coverage:** Each table covers three periods:
   - 前々期 (Two Periods Ago)
   - 前期 (Previous Period)
   - 当期 (Current Period)

4. **Data Types per Period:**
   - 前々期: 2 columns (金額, 売上比)
   - 前期: 4 columns (金額, 売上比, 前年比, 対年変更)
   - 当期: 4 columns (金額, 売上比, 前年比, 対年変更)

5. **Merged Cells:** Table titles use merged cells spanning multiple columns and rows

6. **Total Rows:** Each table has clear subtotal and/or total rows with specific formulas

---

## Usage Notes for Programming

### Reading Data
- Use row numbers as specified in each table structure
- Item names are always in Column A (Tables ①③④) or Column M (Tables ②⑤)
- Financial data starts from the columns specified in the column mapping

### Writing Data
- Ensure formulas in 売上比, 前年比, and 対年変更 columns are preserved
- Total rows contain SUM formulas that should be maintained
- Empty item rows (e.g., Row 7, 8 in Table ①) can be used for additional data

### Cell References
- Most formulas reference $B$9 (Table ① total) as the base for percentage calculations
- Subtotals use SUM formulas over specific ranges
- YoY calculations use IFERROR to handle division by zero

---

## File Location
**Full Path:** c:\Users\t.nohora\OneDrive - 御堂筋税理士法人\デスクトップ\MyGasProject\経営の診断書_フォーマット.xlsx

**Analysis Date:** 2025-10-26
