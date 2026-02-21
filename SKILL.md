---
name: excel-sql-skill
description: "Query and edit live Excel files using SQL via xlwings + pandasql. Use when: working with large Excel datasets that need SQL-style queries, game balancing data, financial models with formulas. Requires Excel to be installed and open."
---

# Excel SQL Skill

Query and edit **live Excel files** using SQL. Uses xlwings (Excel COM/AppleScript bridge) + pandasql (in-memory SQL on DataFrames).

## Prerequisites

- Excel must be **installed and open** on the target machine
- Python 3.9+ with dependencies: `pip install -r requirements.txt`
- Windows: full live editing support (COM automation)
- Mac: works but may need manual save trigger

## Quick Start

```bash
# 1. Install dependencies
pip install xlwings pandas pandasql

# 2. Run the helper script
python scripts/excel_sql.py
```

## Usage from Claude Code

### Attach to an open Excel workbook

```bash
# Default: row 3 is the header (rows 1-2 are type/meta info — common in game data sheets)
python scripts/excel_sql.py attach "Book1.xlsx"

# Attach to the active workbook
python scripts/excel_sql.py attach

# Override header row if your sheet is different
python scripts/excel_sql.py attach "Book1.xlsx" --header-row 1

# Reload with a different header row
python scripts/excel_sql.py reload --header-row 2
```

> **Note:** The default header row is **3**. This matches the common game data sheet format where rows 1–2 contain type annotations and metadata, and row 3 has the actual column names. Pass `--header-row N` to override.

### Query a sheet with SQL

```python
python scripts/excel_sql.py query "SELECT * FROM Sheet1 WHERE level > 10"
```

Sheet names become table names. Spaces in sheet names → use double quotes:
```python
python scripts/excel_sql.py query 'SELECT * FROM "Monster Data" WHERE hp > 100'
```

### Write query results back to Excel

```python
# Update cells matching a condition
python scripts/excel_sql.py exec "UPDATE Sheet1 SET damage = damage * 1.2 WHERE class = 'warrior'"

# Insert new rows
python scripts/excel_sql.py exec "INSERT INTO Sheet1 (name, level, hp) VALUES ('Dragon', 50, 9999)"
```

### Reload sheets (after external edits)

```python
python scripts/excel_sql.py reload
```

## Workflow

```
[Initial Setup]
1. Planner opens Excel file with balancing data
2. Claude Code attaches via: python scripts/excel_sql.py attach "game_balance.xlsx"

[Query Loop]
3. User: "Show me all monsters with HP > 1000"
4. Claude Code: python scripts/excel_sql.py query "SELECT name, hp, damage FROM Monsters WHERE hp > 1000"
5. User: "Increase all boss damage by 15%"
6. Claude Code: python scripts/excel_sql.py exec "UPDATE Monsters SET damage = ROUND(damage * 1.15) WHERE is_boss = 1"
7. Excel auto-recalculates formulas → Planner sees results immediately
8. Repeat...
```

## How It Works

1. **xlwings** attaches to a running Excel process (no file locking issues)
2. Sheets are loaded into **pandas DataFrames** in memory
3. **pandasql** runs SQL queries against those DataFrames
4. For writes: changes are diffed and written back cell-by-cell via xlwings
5. Excel's formula engine recalculates automatically on write

## Notes

- No intermediate SQLite files — everything stays in memory
- Formula cells are preserved; only value cells are overwritten
- Column headers come from row 1 of each sheet
- Large sheets (>100k rows) may be slow on initial load
