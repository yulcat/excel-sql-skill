---
name: excel-sql-skill
description: "Query and edit live Excel files using SQL via xlwings + pandasql. Use when: working with large Excel datasets that need SQL-style queries, game balancing data, financial models with formulas. Requires Excel to be installed and open."
---

# Excel SQL Skill

Query and edit **live Excel files** using SQL. Uses xlwings (Excel COM/AppleScript bridge) + pandasql (in-memory SQL on DataFrames).

## Prerequisites

- Excel installed and open on the target machine
- Python 3.9+: `pip install -r requirements.txt`
- Windows: full live editing (COM automation)
- Mac: works, may need manual save trigger

## Workflow

```
[Setup]    python scripts/excel_sql.py attach "file.xlsx"
[Query]    python scripts/excel_sql.py query "SELECT * FROM Sheet1 WHERE hp > 100"
[Edit]     python scripts/excel_sql.py exec "UPDATE Sheet1 SET damage = damage * 1.2 WHERE is_boss = 1"
[Reload]   python scripts/excel_sql.py reload
```

Excel auto-recalculates formulas on every write. The user sees changes immediately.

## Commands

```bash
attach [workbook] [--header-row N]   # Attach to open workbook (default header row: 3)
reload [--header-row N]              # Reload sheets from attached workbook
query "SQL"                          # SELECT — prints results
exec "SQL"                           # UPDATE / INSERT / DELETE — writes back to Excel
sheets                               # List loaded sheets and row counts
schema [sheet]                       # Show column names and types
```

**Header row default is 3.** Rows 1–2 are treated as type/meta annotations (common in game data sheets). Override with `--header-row 1` if your sheet has headers in row 1.

Sheet names become SQL table names. Use double quotes for spaces: `"Monster Data"`.

## How It Works

1. xlwings attaches to the running Excel process
2. Sheets load into pandas DataFrames in memory
3. pandasql runs SQL against those DataFrames
4. Writes diff only changed cells back via xlwings → formula recalculation triggers automatically

## Working with Large Datasets

For sheets with hundreds or thousands of rows, follow this order:
1. **Sample** reference sheets to understand patterns (`LIMIT 20`, `GROUP BY` aggregations)
2. **Propose rules** to the user based on observed patterns — get explicit confirmation
3. **Apply rules** via SQL UPDATE (avoids loading bulk data into context)
4. **Batch** (50–100 rows at a time) only if row-by-row judgment is truly unavoidable

See [references/workflow.md](references/workflow.md) for the full guide with examples.

## Examples

See [references/examples.md](references/examples.md) for common patterns including game balancing workflows, schema inspection, and bulk updates.
