import argparse
import re
import sys
from pathlib import Path

import pandas as pd
from openpyxl.styles import Font, PatternFill, Alignment
from openpyxl.utils import get_column_letter

SUPPORTED = {".xlsx", ".xls", ".csv"}

ALIASES = {
    "email": ["e mail", "email address", "mail", "e mail address"],
    "phone": ["phone number", "telephone", "tel", "mobile", "phone no"],
    "name": ["full name", "customer name", "client name"],
    "company": ["company name", "organization", "organisation", "firm"],
    "city": ["town"],
    "amount": ["total", "total amount", "price", "value"],
    "date": ["order date", "created", "created at"],
}


def normalize_header(text):
    text = str(text).strip().lower()
    text = re.sub(r"[_\-\.:/]+", " ", text)
    text = re.sub(r"[^\w\s]", "", text)
    text = re.sub(r"\s+", " ", text).strip()
    for canonical, variants in ALIASES.items():
        if text == canonical or text in variants:
            return canonical
    return text.replace(" ", "_")


def read_file(path):
    if path.suffix.lower() == ".csv":
        return {path.name: pd.read_csv(path, dtype=str, keep_default_na=False)}
    sheets = pd.read_excel(path, sheet_name=None, dtype=str, keep_default_na=False)
    if len(sheets) == 1:
        return {path.name: next(iter(sheets.values()))}
    return {f"{path.name} [{name}]": df for name, df in sheets.items()}


def clean(df):
    df = df.copy()
    df.columns = [normalize_header(c) for c in df.columns]
    df = df.loc[:, [c for c in df.columns if not c.startswith("unnamed")]]
    df = df.loc[:, ~df.columns.duplicated()]
    df = df.apply(lambda col: col.astype(str).str.strip())
    df = df.replace({"": pd.NA, "nan": pd.NA, "None": pd.NA})
    df = df.dropna(how="all")
    if "email" in df.columns:
        df["email"] = df["email"].str.lower()
    return df


def style_sheet(ws):
    header_fill = PatternFill("solid", fgColor="1F4E78")
    for cell in ws[1]:
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    ws.freeze_panes = "A2"
    ws.auto_filter.ref = ws.dimensions
    for idx, column in enumerate(ws.columns, start=1):
        longest = max((len(str(c.value)) for c in column if c.value is not None), default=8)
        ws.column_dimensions[get_column_letter(idx)].width = min(max(longest + 2, 10), 50)


def merge(input_dir, output_file, dedupe_on=None):
    files = sorted(p for p in Path(input_dir).iterdir() if p.suffix.lower() in SUPPORTED and not p.name.startswith("~$"))
    if not files:
        sys.exit(f"No Excel/CSV files found in {input_dir}")

    frames, report = [], []
    for path in files:
        try:
            parts = read_file(path)
        except Exception as exc:
            report.append({"source": path.name, "rows_read": 0, "status": f"skipped: {exc}"})
            continue
        for label, raw in parts.items():
            df = clean(raw)
            df.insert(0, "source_file", label)
            frames.append(df)
            report.append({"source": label, "rows_read": len(df), "status": "ok"})

    if not frames:
        sys.exit("No readable data found.")

    merged = pd.concat(frames, ignore_index=True, sort=False)
    before = len(merged)
    keys = [k for k in (dedupe_on or []) if k in merged.columns]
    subset = keys or [c for c in merged.columns if c != "source_file"]
    merged = merged.drop_duplicates(subset=subset, keep="first").reset_index(drop=True)
    removed = before - len(merged)

    summary = pd.DataFrame(report)
    totals = pd.DataFrame([
        {"source": "TOTAL rows read", "rows_read": before, "status": ""},
        {"source": "Duplicates removed", "rows_read": removed, "status": f"key: {', '.join(subset)}"},
        {"source": "Rows in output", "rows_read": len(merged), "status": ""},
    ])
    summary = pd.concat([summary, totals], ignore_index=True)

    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        merged.to_excel(writer, sheet_name="Merged", index=False)
        summary.to_excel(writer, sheet_name="Summary", index=False)
        for ws in writer.book.worksheets:
            style_sheet(ws)

    return len(files), before, removed, len(merged)


def main():
    parser = argparse.ArgumentParser(description="Merge and clean multiple Excel/CSV files into one workbook.")
    parser.add_argument("input_dir", help="Folder containing .xlsx, .xls or .csv files")
    parser.add_argument("-o", "--output", default="merged.xlsx", help="Output Excel file (default: merged.xlsx)")
    parser.add_argument("-k", "--key", nargs="*", help="Column(s) used to detect duplicates, e.g. -k email")
    args = parser.parse_args()

    n_files, read, removed, kept = merge(args.input_dir, args.output, args.key)
    print(f"Files: {n_files} | Rows read: {read} | Duplicates removed: {removed} | Rows written: {kept}")
    print(f"Saved -> {args.output}")


if __name__ == "__main__":
    main()
