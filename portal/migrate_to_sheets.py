"""
migrate_to_sheets.py — one-time upload of all local M2 CSVs and M3
reference files into Google Sheets.

Run once from the project root:
    cd "c:/PowerUpInfinite/Pre-Onboarding Portal"
    python portal/migrate_to_sheets.py

After this runs successfully, the local CSV/Excel files are no longer
needed for the portal — Google Sheets is the source of truth.

Safe to re-run: uses upsert logic for client data (dedup by primary key),
and full-replace for reference data (M3 + Scheme_Category).
"""

import os
import sys
sys.stdout.reconfigure(encoding="utf-8", errors="replace")
sys.stderr.reconfigure(encoding="utf-8", errors="replace")
sys.path.insert(0, os.path.dirname(__file__))

from dotenv import load_dotenv
load_dotenv(os.path.join(os.path.dirname(__file__), ".env"))

import pandas as pd
import sheets

OK  = "  ✓"
ERR = "  ✗"
SKP = "  -"

# ── Paths to local source files ───────────────────────────────
M2_DIR = os.path.normpath(
    os.path.join(os.path.dirname(__file__), "..", "M2-automation")
)
M3_DIR = os.path.normpath(
    os.path.join(os.path.dirname(__file__), "..", "M3-automation")
)


def _load_csv(path: str) -> pd.DataFrame:
    return pd.read_csv(path, low_memory=False, encoding="utf-8-sig")


def _load_excel(path: str, sheet: int | str = 0) -> pd.DataFrame:
    return pd.read_excel(path, sheet_name=sheet)


def migrate_file(label: str, load_fn, write_fn, key_cols):
    """
    Run one migration step.
    load_fn  — callable returning a DataFrame
    write_fn — sheets.upsert_* or sheets.write_* function
    key_cols — list of key columns (None = full replace)
    """
    print(f"\n  {label}")
    try:
        df = load_fn()
        # Strip BOM from column names if present
        df.columns = [c.lstrip("\ufeff").strip() for c in df.columns]
        print(f"    Loaded: {len(df)} rows, {len(df.columns)} columns")

        result = write_fn(df)

        if isinstance(result, dict):
            print(f"{OK}  Replaced {result['replaced']} rows, added {result['added']} rows "
                  f"({result['total']} total in sheet)")
        else:
            print(f"{OK}  Written {result} rows (full replace)")
    except FileNotFoundError as e:
        print(f"{ERR}  File not found — {e}")
    except Exception as e:
        print(f"{ERR}  Failed — {e}")
        raise


def main():
    print("=" * 60)
    print("PowerUp Portal — Migrate Local Data to Google Sheets")
    print("=" * 60)

    # ── M2 Client Data (upsert by primary key) ────────────────
    print("\n[1] M2 Client Data → Main Spreadsheet")

    migrate_file(
        "PF_level.csv → PF_level",
        lambda: _load_csv(os.path.join(M2_DIR, "PF_level.csv")),
        sheets.upsert_pf_level,
        ["PF_ID"],
    )
    migrate_file(
        "Scheme_level.csv → Scheme_level",
        lambda: _load_csv(os.path.join(M2_DIR, "Scheme_level.csv")),
        sheets.upsert_scheme_level,
        ["PF_ID", "ISIN"],
    )
    migrate_file(
        "Riskgroup_level.csv → Riskgroup_level",
        lambda: _load_csv(os.path.join(M2_DIR, "Riskgroup_level.csv")),
        sheets.upsert_riskgroup_level,
        ["PF_ID", "RISK_GROUP_L0"],
    )
    migrate_file(
        "Results.csv → Results",
        lambda: _load_csv(os.path.join(M2_DIR, "Results.csv")),
        sheets.upsert_results,
        ["PF_ID", "TYPE"],
    )
    migrate_file(
        "Lines.csv → Time Series / Lines  (large file — may take a few minutes)",
        lambda: _load_csv(os.path.join(M2_DIR, "Lines.csv")),
        sheets.upsert_lines,
        ["PF_ID", "DATE", "TYPE"],
    )
    migrate_file(
        "Invested_Value_Line.csv → Time Series / Invested_Value_Line",
        lambda: _load_csv(os.path.join(M2_DIR, "Invested_Value_Line.csv")),
        sheets.upsert_invested_value_line,
        ["PF_ID", "DATE"],
    )

    # ── M2 Reference Data (full replace) ─────────────────────
    print("\n[2] M2 Reference Data → Main Spreadsheet")

    migrate_file(
        "Scheme_Category_Catgorization.xlsx → Scheme_Category",
        lambda: _load_excel(os.path.join(M2_DIR, "Scheme_Category_Catgorization.xlsx")),
        sheets.upsert_scheme_category,
        ["Powerup Broad Category"],
    )

    # ── M3 Reference Data (full replace, monthly) ─────────────
    print("\n[3] M3 Reference Data → M3 Reference Spreadsheet")

    migrate_file(
        "AUM_31Jan.csv → AUM",
        lambda: _load_csv(os.path.join(M3_DIR, "AUM_31Jan.csv")),
        sheets.write_m3_aum,
        None,
    )
    migrate_file(
        "Powerranking.csv → Powerranking",
        lambda: _load_csv(os.path.join(M3_DIR, "Powerranking.csv")),
        sheets.write_m3_powerranking,
        None,
    )
    migrate_file(
        "upside_downside_mar.xlsx → Upside_Downside",
        lambda: _load_excel(os.path.join(M3_DIR, "upside_downside_mar.xlsx")),
        sheets.write_m3_upside_downside,
        None,
    )
    migrate_file(
        "Rolling_Returns_Mar.csv → Rolling_Returns",
        lambda: _load_csv(os.path.join(M3_DIR, "Rolling_Returns_Mar.csv")),
        sheets.write_m3_rolling_returns,
        None,
    )

    print("\n" + "=" * 60)
    print("Migration complete.")
    print("Check any ✗ lines above — those sheets were not written.")
    print("=" * 60)


if __name__ == "__main__":
    main()
