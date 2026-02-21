#!/usr/bin/env python3
"""
excel_sql.py — Query and edit live Excel files using SQL.

Uses xlwings (Excel process bridge) + pandasql (in-memory SQL on DataFrames).
Designed as a CLI helper for Claude Code / OpenClaw.

Usage:
    python excel_sql.py attach [workbook_name]   # Attach to open Excel workbook
    python excel_sql.py reload                     # Reload all sheets from Excel
    python excel_sql.py query "SQL"                # SELECT query, prints results
    python excel_sql.py exec "SQL"                 # UPDATE/INSERT, writes back to Excel
    python excel_sql.py sheets                     # List available sheets/tables
    python excel_sql.py schema [sheet]             # Show column info for a sheet
"""

import sys
import json
import re
import argparse
from pathlib import Path

import pandas as pd
import pandasql as psql

try:
    import xlwings as xw
except ImportError:
    xw = None

# ---------------------------------------------------------------------------
# Global state (persisted in a temp JSON between calls)
# ---------------------------------------------------------------------------
STATE_FILE = Path.home() / ".excel_sql_state.json"

_wb = None          # xlwings Workbook reference
_frames = {}        # sheet_name -> DataFrame


def _save_state(wb_name: str):
    """Save minimal state so subsequent CLI calls can re-attach."""
    STATE_FILE.write_text(json.dumps({"workbook": wb_name}))


def _load_state() -> dict:
    if STATE_FILE.exists():
        return json.loads(STATE_FILE.read_text())
    return {}


# ---------------------------------------------------------------------------
# Core functions
# ---------------------------------------------------------------------------

def attach(workbook_name: str | None = None) -> str:
    """Attach to an open Excel workbook and load all sheets."""
    global _wb, _frames

    if xw is None:
        return "ERROR: xlwings not installed. Run: pip install xlwings"

    try:
        if workbook_name:
            _wb = xw.Book(workbook_name)
        else:
            _wb = xw.books.active
    except Exception as e:
        return f"ERROR: Could not attach to workbook: {e}"

    if _wb is None:
        return "ERROR: No active workbook found. Is Excel open?"

    _frames.clear()
    for sheet in _wb.sheets:
        name = sheet.name
        # Read used range into DataFrame
        data = sheet.used_range.options(pd.DataFrame, header=True, index=False).value
        if data is not None and not data.empty:
            # Sanitize column names
            data.columns = [str(c).strip() for c in data.columns]
            _frames[name] = data

    _save_state(_wb.name)
    sheets_info = ", ".join(f"{k} ({len(v)} rows)" for k, v in _frames.items())
    return f"Attached to '{_wb.name}'. Sheets: {sheets_info}"


def reload() -> str:
    """Reload sheets from the currently attached workbook."""
    state = _load_state()
    wb_name = state.get("workbook")
    if not wb_name:
        return "ERROR: No workbook attached. Run: attach [name]"
    return attach(wb_name)


def list_sheets() -> str:
    """List loaded sheets and their row counts."""
    _ensure_loaded()
    if not _frames:
        return "No sheets loaded. Run: attach [workbook]"
    lines = []
    for name, df in _frames.items():
        lines.append(f"  {name}: {len(df)} rows, {len(df.columns)} columns")
    return "Loaded sheets:\n" + "\n".join(lines)


def schema(sheet_name: str | None = None) -> str:
    """Show column names and dtypes for a sheet (or all sheets)."""
    _ensure_loaded()
    targets = [sheet_name] if sheet_name else list(_frames.keys())
    lines = []
    for name in targets:
        df = _frames.get(name)
        if df is None:
            lines.append(f"{name}: NOT FOUND")
            continue
        cols = ", ".join(f"{c} ({df[c].dtype})" for c in df.columns)
        lines.append(f"{name}: {cols}")
    return "\n".join(lines)


def query(sql: str) -> str:
    """Run a SELECT query and return results as formatted text."""
    _ensure_loaded()
    env = _build_env()
    try:
        result = psql.sqldf(sql, env)
    except Exception as e:
        return f"SQL ERROR: {e}"

    if result is None or result.empty:
        return "(no results)"

    return result.to_string(index=False)


def exec_sql(sql: str) -> str:
    """
    Run an UPDATE or INSERT statement.
    
    Strategy:
    - Parse the target table name from SQL
    - For UPDATE: run query to identify affected rows, apply changes, write back
    - For INSERT: append rows to DataFrame and write to Excel
    """
    _ensure_loaded()
    sql_upper = sql.strip().upper()

    if sql_upper.startswith("UPDATE"):
        return _handle_update(sql)
    elif sql_upper.startswith("INSERT"):
        return _handle_insert(sql)
    elif sql_upper.startswith("DELETE"):
        return _handle_delete(sql)
    else:
        return f"ERROR: Unsupported statement. Use UPDATE, INSERT, or DELETE."


# ---------------------------------------------------------------------------
# Write-back helpers
# ---------------------------------------------------------------------------

def _handle_update(sql: str) -> str:
    """Handle UPDATE by re-querying and diffing."""
    # Extract table name: UPDATE <table> SET ...
    m = re.match(r"UPDATE\s+[\"']?(\w+)[\"']?\s+SET\s+", sql, re.IGNORECASE)
    if not m:
        return "ERROR: Could not parse UPDATE statement."
    
    table = m.group(1)
    df = _frames.get(table)
    if df is None:
        return f"ERROR: Sheet '{table}' not found."

    # Use pandasql with a trick: create a new version via SELECT
    # We'll parse SET clause and WHERE clause to build the update
    env = _build_env()
    
    # Get the old data
    old_df = df.copy()
    
    # Execute via SQLite (pandasql uses SQLite under the hood)
    # Create table, run update, read back
    import sqlite3
    conn = sqlite3.connect(":memory:")
    
    # Write all frames to SQLite
    for name, frame in _frames.items():
        frame.to_sql(name, conn, index=False, if_exists="replace")
    
    # Run the UPDATE
    cursor = conn.cursor()
    try:
        cursor.execute(sql)
        affected = cursor.rowcount
        conn.commit()
    except Exception as e:
        conn.close()
        return f"SQL ERROR: {e}"
    
    # Read back the updated table
    new_df = pd.read_sql(f'SELECT * FROM "{table}"', conn)
    conn.close()
    
    # Diff and write changes to Excel
    changes = _diff_and_write(table, old_df, new_df)
    _frames[table] = new_df
    
    return f"Updated {affected} row(s) in '{table}'. {changes} cell(s) written to Excel."


def _handle_insert(sql: str) -> str:
    """Handle INSERT by executing in SQLite and appending to Excel."""
    m = re.match(r"INSERT\s+INTO\s+[\"']?(\w+)[\"']?", sql, re.IGNORECASE)
    if not m:
        return "ERROR: Could not parse INSERT statement."
    
    table = m.group(1)
    df = _frames.get(table)
    if df is None:
        return f"ERROR: Sheet '{table}' not found."

    import sqlite3
    conn = sqlite3.connect(":memory:")
    for name, frame in _frames.items():
        frame.to_sql(name, conn, index=False, if_exists="replace")
    
    cursor = conn.cursor()
    try:
        cursor.execute(sql)
        affected = cursor.rowcount
        conn.commit()
    except Exception as e:
        conn.close()
        return f"SQL ERROR: {e}"
    
    new_df = pd.read_sql(f'SELECT * FROM "{table}"', conn)
    conn.close()
    
    # Write new rows to Excel
    _append_rows(table, df, new_df)
    _frames[table] = new_df
    
    return f"Inserted {affected} row(s) into '{table}'."


def _handle_delete(sql: str) -> str:
    """Handle DELETE by executing in SQLite and rewriting sheet."""
    m = re.match(r"DELETE\s+FROM\s+[\"']?(\w+)[\"']?", sql, re.IGNORECASE)
    if not m:
        return "ERROR: Could not parse DELETE statement."
    
    table = m.group(1)
    df = _frames.get(table)
    if df is None:
        return f"ERROR: Sheet '{table}' not found."

    import sqlite3
    conn = sqlite3.connect(":memory:")
    for name, frame in _frames.items():
        frame.to_sql(name, conn, index=False, if_exists="replace")
    
    cursor = conn.cursor()
    try:
        cursor.execute(sql)
        affected = cursor.rowcount
        conn.commit()
    except Exception as e:
        conn.close()
        return f"SQL ERROR: {e}"
    
    new_df = pd.read_sql(f'SELECT * FROM "{table}"', conn)
    conn.close()
    
    # Rewrite entire sheet
    _rewrite_sheet(table, new_df)
    _frames[table] = new_df
    
    return f"Deleted {affected} row(s) from '{table}'."


# ---------------------------------------------------------------------------
# Excel write helpers
# ---------------------------------------------------------------------------

def _diff_and_write(sheet_name: str, old_df: pd.DataFrame, new_df: pd.DataFrame) -> int:
    """Write only changed cells back to Excel. Returns number of cells written."""
    if _wb is None:
        return 0
    
    sheet = _wb.sheets[sheet_name]
    changes = 0
    
    for i in range(min(len(old_df), len(new_df))):
        for j, col in enumerate(new_df.columns):
            old_val = old_df.iloc[i, j] if j < len(old_df.columns) else None
            new_val = new_df.iloc[i, j]
            if not _values_equal(old_val, new_val):
                # +1 for header row, +1 for 1-indexed
                sheet.range((i + 2, j + 1)).value = new_val
                changes += 1
    
    return changes


def _append_rows(sheet_name: str, old_df: pd.DataFrame, new_df: pd.DataFrame):
    """Append new rows to the Excel sheet."""
    if _wb is None:
        return
    
    sheet = _wb.sheets[sheet_name]
    start_row = len(old_df) + 2  # +1 header, +1 for 1-indexed
    
    for i in range(len(old_df), len(new_df)):
        for j, col in enumerate(new_df.columns):
            sheet.range((start_row + (i - len(old_df)), j + 1)).value = new_df.iloc[i, j]


def _rewrite_sheet(sheet_name: str, new_df: pd.DataFrame):
    """Clear and rewrite an entire sheet (for DELETE operations)."""
    if _wb is None:
        return
    
    sheet = _wb.sheets[sheet_name]
    sheet.used_range.clear_contents()
    
    # Write header
    for j, col in enumerate(new_df.columns):
        sheet.range((1, j + 1)).value = col
    
    # Write data
    if not new_df.empty:
        sheet.range((2, 1)).value = new_df.values.tolist()


def _values_equal(a, b) -> bool:
    """Compare two values, handling NaN."""
    if pd.isna(a) and pd.isna(b):
        return True
    try:
        return a == b
    except (TypeError, ValueError):
        return str(a) == str(b)


# ---------------------------------------------------------------------------
# Helpers
# ---------------------------------------------------------------------------

def _ensure_loaded():
    """Auto-reload if frames are empty (fresh CLI invocation)."""
    global _frames
    if not _frames:
        state = _load_state()
        wb_name = state.get("workbook")
        if wb_name and xw is not None:
            try:
                attach(wb_name)
            except Exception:
                pass


def _build_env() -> dict:
    """Build environment dict for pandasql (sheet_name -> DataFrame)."""
    return dict(_frames)


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------

def main():
    parser = argparse.ArgumentParser(description="Excel SQL helper for OpenClaw")
    sub = parser.add_subparsers(dest="command")

    p_attach = sub.add_parser("attach", help="Attach to an open Excel workbook")
    p_attach.add_argument("workbook", nargs="?", default=None)

    sub.add_parser("reload", help="Reload sheets from attached workbook")
    sub.add_parser("sheets", help="List loaded sheets")

    p_schema = sub.add_parser("schema", help="Show schema for sheet(s)")
    p_schema.add_argument("sheet", nargs="?", default=None)

    p_query = sub.add_parser("query", help="Run SELECT query")
    p_query.add_argument("sql")

    p_exec = sub.add_parser("exec", help="Run UPDATE/INSERT/DELETE")
    p_exec.add_argument("sql")

    args = parser.parse_args()

    if args.command == "attach":
        print(attach(args.workbook))
    elif args.command == "reload":
        print(reload())
    elif args.command == "sheets":
        print(list_sheets())
    elif args.command == "schema":
        print(schema(args.sheet))
    elif args.command == "query":
        print(query(args.sql))
    elif args.command == "exec":
        print(exec_sql(args.sql))
    else:
        parser.print_help()


if __name__ == "__main__":
    main()
