# Excel Merger

Merges and cleans multiple Excel / CSV files into one tidy workbook — in seconds.

Typical use: monthly sales sheets, contact lists or exports from different systems that all have *slightly different* column names, empty rows and duplicate records.

## What it does

- Reads every `.xlsx`, `.xls` and `.csv` file in a folder (all sheets)
- **Unifies column names** — `E-mail`, `email address`, `Email` → `email`; `Tel`, `Mobile`, `Phone Number` → `phone`
- Removes empty rows and stray spaces, lower-cases emails
- **Removes duplicates** (by all columns, or by a key you choose, e.g. `email`)
- Adds a `source_file` column so every row can be traced back
- Writes a formatted workbook: bold header, frozen top row, filters, auto column widths
- Adds a **Summary** sheet: rows per file, duplicates removed, final count

## Example

Input (3 files, 3 different header styles):

| january_sales.xlsx | february_sales.xlsx | march_sales.csv |
|---|---|---|
| Full Name, E-mail, Phone Number, Total | customer name, email address, Tel, Price | Name, Email, Mobile, Amount, City |

```
python merge.py sample_data -o merged.xlsx -k email
Files: 3 | Rows read: 7 | Duplicates removed: 2 | Rows written: 5
```

Output — one clean table:

| source_file | name | email | phone | amount | city |
|---|---|---|---|---|---|
| february_sales.xlsx | John Doe | john@example.com | 555-0102 | 85 | |
| february_sales.xlsx | Lucas Brown | lucas@example.com | 555-0104 | 210 | |
| january_sales.xlsx | Anna Smith | anna@example.com | 555-0101 | 120 | |
| january_sales.xlsx | Mary Jones | mary@example.com | 555-0103 | 40 | |
| march_sales.csv | Emma White | emma@example.com | 555-0105 | 99 | Boston |

## Usage

```bash
pip install -r requirements.txt
python merge.py <folder> -o merged.xlsx            # dedupe on all columns
python merge.py <folder> -o merged.xlsx -k email   # dedupe on email only
```

## Customising

Column name mappings live in the `ALIASES` dictionary at the top of `merge.py` — add your own synonyms in one line.

## Need something similar?

I build custom Python automations for Excel, reports and data cleaning. Feel free to reach out.
