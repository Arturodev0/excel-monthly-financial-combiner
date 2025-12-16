import os
import pandas as pd
import re
import warnings
import argparse

warnings.filterwarnings("ignore")

PL_SHEETS_TRY_DEFAULT = ["P&L", "P&L by Month"]
BS_SHEET_DEFAULT = "BS by Month Condensed"
DB_SHEET_DEFAULT = "DataBase Result"

OUTPUT_DEFAULT = "combined.xlsx"
MONTHLY_DEFAULT = "monthly.xlsx"


def _is_worksheet_not_found(err: Exception) -> bool:
    msg = str(err)
    return isinstance(err, ValueError) and "Worksheet named" in msg and "not found" in msg


def _available_sheets(file_path: str) -> list[str]:
    """
    Lista hojas SOLO para mensajes de error (no afecta lógica).
    """
    try:
        import openpyxl
        wb = openpyxl.load_workbook(file_path, read_only=True, data_only=True)
        return list(wb.sheetnames)
    except Exception:
        return []


def _err_ctx(source: str, file_path: str) -> str:
    return f"Source={source} | File={file_path}"


def _normalize_source_in_combined(s: str) -> str:
    s = str(s).strip()
    if "/" in s:
        return s
    parts = s.split(".")
    if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
        return f"{int(parts[0]):02d}.{parts[1]}"
    return s


def parse_args():
    p = argparse.ArgumentParser(prog="combine_monthly_financials.py")
    p.add_argument("--base-dir", default=None)
    p.add_argument("--output", default=OUTPUT_DEFAULT)
    p.add_argument("--monthly", default=MONTHLY_DEFAULT)
    p.add_argument("--pl-sheets", nargs="*", default=PL_SHEETS_TRY_DEFAULT)
    p.add_argument("--bs-sheet", default=BS_SHEET_DEFAULT)
    p.add_argument("--db-sheet", default=DB_SHEET_DEFAULT)
    return p.parse_args()


def main():
    args = parse_args()

    current_dir = os.path.dirname(os.path.abspath(__file__))
    base_dir = os.path.abspath(args.base_dir) if args.base_dir else current_dir

    output_file = os.path.join(base_dir, args.output)
    pl_data = []
    bs_data = []
    db_data = []
    existing_sources = set()

    if os.path.exists(output_file):
        try:
            existing_df = pd.read_excel(output_file, sheet_name="P&L Combined", engine="openpyxl")
            if "Source" in existing_df.columns:
                existing_df["Source"] = existing_df["Source"].astype(str).apply(_normalize_source_in_combined)
                existing_sources = set(existing_df["Source"].dropna().unique())
                print(f"✅ Loaded existing P&L Combined with {len(existing_sources)} sources.")
            else:
                print("⚠️ Existing 'P&L Combined' found but column 'Source' is missing; duplicates may occur.")
        except Exception as e:
            print(f"❌ Failed to read existing file '{output_file}' (sheet 'P&L Combined'). "
                  f"Type={type(e).__name__} | Error={e}")

    for year_folder in os.listdir(base_dir):
        year_path = os.path.join(base_dir, year_folder)
        if os.path.isdir(year_path) and year_folder.isdigit():
            for folder_name in os.listdir(year_path):
                folder_path = os.path.join(year_path, folder_name)
                if os.path.isdir(folder_path) and '.' in folder_name:
                    parts = folder_name.strip().split(".")
                    if len(parts) == 2 and parts[0].isdigit() and parts[1].isdigit():
                        month = f"{int(parts[0]):02d}"
                        year = parts[1]
                        normalized_source = f"{year_folder}/{month}.{year}"
                        file_date = f"{year}-{month}-01"

                        if normalized_source in existing_sources:
                            continue

                        file_path = os.path.join(folder_path, args.monthly)
                        if os.path.exists(file_path):
                            try:
                                try:
                                    pl_df = None
                                    pl_sheet_used = None

                                    for sheet_try in args.pl_sheets:
                                        try:
                                            pl_df = pd.read_excel(file_path, sheet_name=sheet_try, engine="openpyxl")
                                            pl_sheet_used = sheet_try
                                            break
                                        except Exception as e:
                                            if _is_worksheet_not_found(e):
                                                continue
                                            raise
                                        
                                    if pl_df is not None and "Amount" not in pl_df.columns:
                                        month_name_map = {
                                            "01": "January", "02": "February", "03": "March", "04": "April",
                                            "05": "May", "06": "June", "07": "July", "08": "August",
                                            "09": "September", "10": "October", "11": "November", "12": "December"
                                        }

                                        month_expected = month_name_map.get(month, "").lower()

                                        banned = {"parent", "category", "total"}
                                        candidates = [c for c in pl_df.columns if str(c).strip().lower() not in banned]

                                        picked = None
                                        for c in candidates:
                                            if str(c).strip().lower() == month_expected:
                                                picked = c
                                                break

                                        if picked is None and len(candidates) == 1:
                                            picked = candidates[0]

                                        if picked is None:
                                            raise ValueError(
                                                f"❌ P&L sin 'Amount' y no pude identificar columna del mes. "
                                                f"Esperaba algo como '{month_name_map.get(month)}'. Columnas: {list(pl_df.columns)}"
                                            )

                                        pl_df["Amount"] = pd.to_numeric(pl_df[picked], errors="coerce")

                                        drop_cols = [picked]
                                        if "Total" in pl_df.columns:
                                            drop_cols.append("Total")
                                        pl_df = pl_df.drop(columns=drop_cols, errors="ignore")

                                        print(f"ℹ️ Normalized P&L by Month: '{picked}' -> 'Amount' | Source={normalized_source}")
                                    

                                    if pl_df is None:
                                        sheets = _available_sheets(file_path)
                                        sheets_msg = f"Available sheets: {sheets}" if sheets else "Available sheets: (could not read)"
                                        print(
                                            "❌ P&L sheet not found. "
                                            f"Tried={args.pl_sheets}. {sheets_msg}. "
                                            f"{_err_ctx(normalized_source, file_path)}"
                                        )
                                    else:
                                        if pl_sheet_used != "P&L":
                                            print(f"ℹ️ P&L loaded using sheet '{pl_sheet_used}'. {_err_ctx(normalized_source, file_path)}")

                                        pl_df["Source"] = normalized_source
                                        pl_df["Date"] = pd.to_datetime(file_date)
                                        pl_df["Week"] = pl_df["Date"].dt.isocalendar().week

                                        required_cols = ["Parent", "Category"]
                                        for col in required_cols:
                                            if col not in pl_df.columns:
                                                raise ValueError(f"❌ Columna faltante en P&L: {col} | {_err_ctx(normalized_source, file_path)}")

                                        mask_total = pl_df["Parent"].notna() & pl_df["Parent"].astype(str).str.strip().ne("")
                                        pl_df.loc[mask_total, "Category"] = "Total " + pl_df.loc[mask_total, "Category"].astype(str)

                                        cols_to_fill = [col for col in pl_df.columns if col not in ["Amount", "Date", "Source"]]
                                        pl_df[cols_to_fill] = pl_df[cols_to_fill].fillna(method="ffill")

                                        pl_df["Parent"] = pl_df["Parent"].astype(str).str.strip().str.title()
                                        parent_map = {
                                            "Income": "1 Income",
                                            "Cogs": "2 COGS",
                                            "Gross Profit": "3 Gross Profit",
                                            "Expenses": "5 Expenses",
                                            "Net Ordinary Income": "6 Net Ordinary Income",
                                            "Other Income": "7 Other Income",
                                            "Other Expenses": "8 Other Expenses",
                                            "Net Income": "9 Net Income"
                                        }
                                        pl_df["Parent"] = pl_df["Parent"].replace(parent_map)

                                        pl_df = pl_df[~pl_df["Parent"].isin(["3 Gross Profit", "6 Net Ordinary Income", "9 Net Income"])]

                                        total_mask = pl_df["Category"].astype(str).str.contains("Total", case=False)
                                        parent_filter = pl_df["Parent"].isin([
                                            "1 Income", "2 COGS", "5 Expenses",
                                            "7 Other Income", "8 Other Expenses",
                                            "Income", "COGS", "Expenses", "Other Income", "Other Expenses"
                                        ])
                                        final_mask = ~(total_mask & parent_filter)
                                        pl_df = pl_df[final_mask]

                                        pl_data.append(pl_df)

                                except Exception as e:
                                    sheets = _available_sheets(file_path) if _is_worksheet_not_found(e) else []
                                    sheets_msg = f" | Available sheets: {sheets}" if sheets else ""
                                    print(
                                        f"❌ Error procesando P&L → Type={type(e).__name__} | Error={e}"
                                        f"{sheets_msg} | {_err_ctx(normalized_source, file_path)}"
                                    )

                                try:
                                    bs_df_raw = pd.read_excel(file_path, sheet_name=args.bs_sheet, engine="openpyxl")

                                    date_cols = [col for col in bs_df_raw.columns if isinstance(col, str) and re.match(r"^\d{4}-\d{2}$", col)]
                                    if not date_cols:
                                        raise ValueError("❌ No se encontraron columnas con formato 'yyyy-mm' en BS by Month Condensed.")

                                    latest_date_col = sorted(date_cols)[-1]
                                    date_value = latest_date_col + "-01"

                                    cols_to_keep = ["Category", "Category2", "Last Category"]
                                    cols_present = [col for col in cols_to_keep if col in bs_df_raw.columns]

                                    bs_df = bs_df_raw[cols_present + [latest_date_col]].copy()
                                    bs_df = bs_df.rename(columns={latest_date_col: "Amount"})
                                    bs_df["Source"] = normalized_source
                                    bs_df["Date"] = date_value

                                    for col in cols_present:
                                        bs_df[col] = bs_df[col].fillna(method="ffill")

                                    mask_total = bs_df[cols_present].astype(str).apply(lambda x: x.str.contains("Total", case=False, na=False)).any(axis=1)
                                    bs_df = bs_df[~mask_total]

                                    bs_data.append(bs_df)

                                except Exception as e:
                                    if _is_worksheet_not_found(e):
                                        sheets = _available_sheets(file_path)
                                        print(
                                            f"⚠️ Missing sheet '{args.bs_sheet}'. "
                                            f"Available sheets: {sheets if sheets else '(could not read)'} | "
                                            f"{_err_ctx(normalized_source, file_path)}"
                                        )
                                    else:
                                        print(
                                            f"⚠️ Error procesando '{args.bs_sheet}' → Type={type(e).__name__} | Error={e} | "
                                            f"{_err_ctx(normalized_source, file_path)}"
                                        )

                                try:
                                    db_df = pd.read_excel(file_path, sheet_name=args.db_sheet, engine="openpyxl")
                                    db_df["Source"] = normalized_source

                                    if "Date" in db_df.columns:
                                        db_df["Date"] = pd.to_datetime(db_df["Date"], errors="coerce")
                                        db_df["Week"] = db_df["Date"].dt.isocalendar().week.astype("Int64").astype(str).str.zfill(2)
                                    else:
                                        db_df["Date"] = pd.NaT
                                        db_df["Week"] = pd.NA

                                    if "Parent" in db_df.columns:
                                        db_df = db_df[~db_df["Parent"].isin(["3 Gross Profit", "6 Net Ordinary Income", "9 Net Income"])]

                                    db_data.append(db_df)

                                except Exception as e:
                                    if _is_worksheet_not_found(e):
                                        sheets = _available_sheets(file_path)
                                        print(
                                            f"⚠️ Missing sheet '{args.db_sheet}'. "
                                            f"Available sheets: {sheets if sheets else '(could not read)'} | "
                                            f"{_err_ctx(normalized_source, file_path)}"
                                        )
                                    else:
                                        print(
                                            f"⚠️ Error procesando '{args.db_sheet}' → Type={type(e).__name__} | Error={e} | "
                                            f"{_err_ctx(normalized_source, file_path)}"
                                        )

                            except Exception as e:
                                print(
                                    f"❌ Error reading file → Type={type(e).__name__} | Error={e} | "
                                    f"{_err_ctx(normalized_source, file_path)}"
                                )

    if pl_data or bs_data or db_data:
        with pd.ExcelWriter(output_file, engine="openpyxl", mode="w") as writer:
            if pl_data:
                pl_combined = pd.concat(pl_data, ignore_index=True)
                pl_combined.to_excel(writer, sheet_name="P&L Combined", index=False)
                print(f"📄 P&L Combined guardado con {len(pl_combined)} registros.")
            if bs_data:
                bs_combined = pd.concat(bs_data, ignore_index=True)
                bs_combined.to_excel(writer, sheet_name="BS Condensed Combined", index=False)
                print(f"📄 BS Condensed Combined guardado con {len(bs_combined)} registros.")
            if db_data:
                db_combined = pd.concat(db_data, ignore_index=True)
                db_combined.to_excel(writer, sheet_name="DataBase Combined", index=False)
                print(f"📄 DataBase Combined guardado con {len(db_combined)} registros.")
    else:
        print("ℹ️ No hay datos nuevos para agregar.")


if __name__ == "__main__":
    main()
