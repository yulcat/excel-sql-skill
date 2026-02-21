# Examples

## Game Balancing Workflow

```
[Setup]
python scripts/excel_sql.py attach "game_balance.xlsx"
# → Attached to 'game_balance.xlsx' (header row=3). Sheets: Monsters (120 rows), Skills (80 rows)

[Query]
python scripts/excel_sql.py query "SELECT name, hp, damage FROM Monsters WHERE is_boss = 1"
# → lists all boss monsters

[Edit]
python scripts/excel_sql.py exec "UPDATE Monsters SET damage = ROUND(damage * 1.15) WHERE is_boss = 1"
# → Updated 8 row(s). Excel recalculates immediately.

[Verify]
python scripts/excel_sql.py query "SELECT name, damage FROM Monsters WHERE is_boss = 1"
```

## Checking Schema First

```bash
python scripts/excel_sql.py schema
# → Monsters: Id (object), Order (float64), hp (float64), damage (float64), is_boss (float64)...
# → Skills: Id (object), name (object), cooldown (float64), ...
```

## Sheet Names with Spaces

```bash
python scripts/excel_sql.py query 'SELECT * FROM "Monster Data" WHERE hp > 1000'
```

## Non-standard Header Row

```bash
# Sheet has title in row 1, empty row 2, headers in row 3 (default)
python scripts/excel_sql.py attach "data.xlsx"

# Sheet has headers directly in row 1
python scripts/excel_sql.py attach "simple.xlsx" --header-row 1
```

## Bulk Update Pattern

```bash
# 1. Preview what will change
python scripts/excel_sql.py query "SELECT name, hp FROM Monsters WHERE type = 'dragon'"

# 2. Apply the change
python scripts/excel_sql.py exec "UPDATE Monsters SET hp = hp * 1.2 WHERE type = 'dragon'"

# 3. Confirm
python scripts/excel_sql.py query "SELECT name, hp FROM Monsters WHERE type = 'dragon'"
```
