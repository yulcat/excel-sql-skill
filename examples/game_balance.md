# Example: Game Balancing with Excel SQL

## Scenario

A game designer maintains monster/item balancing data in `game_balance.xlsx` with sheets:
- **Monsters**: name, level, hp, attack, defense, is_boss, drop_table
- **Items**: name, type, rarity, price, effect, magnitude
- **DropTable**: monster_name, item_name, drop_rate

## Session

```bash
# 1. Attach to the open workbook
python scripts/excel_sql.py attach "game_balance.xlsx"
# → Attached to 'game_balance.xlsx'. Sheets: Monsters (45 rows), Items (120 rows), DropTable (200 rows)

# 2. Check all bosses
python scripts/excel_sql.py query "SELECT name, level, hp, attack FROM Monsters WHERE is_boss = 1 ORDER BY level"

# 3. Find overpowered monsters (high attack relative to level)
python scripts/excel_sql.py query "SELECT name, level, attack, ROUND(CAST(attack AS FLOAT)/level, 1) as atk_per_lvl FROM Monsters ORDER BY atk_per_lvl DESC LIMIT 10"

# 4. Nerf all boss HP by 10%
python scripts/excel_sql.py exec "UPDATE Monsters SET hp = ROUND(hp * 0.9) WHERE is_boss = 1"
# → Updated 5 row(s) in 'Monsters'. 5 cell(s) written to Excel.
# Designer sees the changes immediately in Excel!

# 5. Cross-sheet query: which items drop from bosses?
python scripts/excel_sql.py query "
    SELECT m.name as monster, i.name as item, i.rarity, d.drop_rate
    FROM Monsters m
    JOIN DropTable d ON m.name = d.monster_name
    JOIN Items i ON d.item_name = i.name
    WHERE m.is_boss = 1
    ORDER BY i.rarity DESC, d.drop_rate DESC
"

# 6. Add a new monster
python scripts/excel_sql.py exec "INSERT INTO Monsters (name, level, hp, attack, defense, is_boss) VALUES ('Shadow Dragon', 55, 12000, 450, 300, 1)"

# 7. Adjust all legendary item prices
python scripts/excel_sql.py exec "UPDATE Items SET price = ROUND(price * 1.25) WHERE rarity = 'legendary'"
```

## Tips

- After major changes, ask the designer to verify formulas are recalculating correctly
- Use `reload` if the designer makes manual edits in Excel
- Sheet names with spaces need double quotes in SQL: `"Monster Data"`
