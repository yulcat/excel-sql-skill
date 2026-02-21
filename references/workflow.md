# Recommended Workflow for Large Datasets

## The Problem

Reading or writing thousands of rows at once bloats the context window and degrades quality.
The goal is to **never load raw data in bulk** — instead, derive rules and apply them via SQL.

## Recommended Approach

### Step 1: Understand the reference sheets

Load a sample from each reference sheet to understand structure and patterns:

```bash
python scripts/excel_sql.py schema
python scripts/excel_sql.py query "SELECT * FROM RefSheet1 LIMIT 20"
python scripts/excel_sql.py query "SELECT * FROM RefSheet2 LIMIT 20"
```

Use aggregations to spot trends without loading every row:

```bash
python scripts/excel_sql.py query "SELECT type, AVG(reward), MIN(reward), MAX(reward) FROM RefSheet1 GROUP BY type"
```

### Step 2: Derive rules with the user

Propose explicit rules based on what you observed. Example:

> "Based on the reference sheet, it looks like:
> - Boss monsters (is_boss=1) with hp > 5000 → reward_tier = 'S'
> - Boss monsters with hp 2000–5000 → reward_tier = 'A'
> - Non-boss with damage > 300 → reward_tier = 'B'
> - Everything else → reward_tier = 'C'
>
> Does this match your intent?"

Get explicit confirmation before writing anything.

### Step 3: Apply rules via SQL UPDATE

Once rules are agreed, apply them in a single UPDATE (or a few targeted ones):

```bash
python scripts/excel_sql.py exec "UPDATE Monsters SET reward_tier = 'S' WHERE is_boss = 1 AND hp > 5000"
python scripts/excel_sql.py exec "UPDATE Monsters SET reward_tier = 'A' WHERE is_boss = 1 AND hp BETWEEN 2000 AND 5000"
python scripts/excel_sql.py exec "UPDATE Monsters SET reward_tier = 'B' WHERE is_boss = 0 AND damage > 300"
python scripts/excel_sql.py exec "UPDATE Monsters SET reward_tier = 'C' WHERE reward_tier IS NULL"
```

### Step 4: Verify with spot checks

```bash
python scripts/excel_sql.py query "SELECT reward_tier, COUNT(*) FROM Monsters GROUP BY reward_tier"
python scripts/excel_sql.py query "SELECT name, hp, damage, reward_tier FROM Monsters WHERE reward_tier = 'S' LIMIT 10"
```

### Step 5: Batch processing (fallback)

If row-by-row judgment is unavoidable (e.g., each row needs unique manual reasoning),
process in batches of 50–100 rows:

```bash
python scripts/excel_sql.py query "SELECT * FROM Monsters LIMIT 100 OFFSET 0"
# → review, apply updates
python scripts/excel_sql.py query "SELECT * FROM Monsters LIMIT 100 OFFSET 100"
# → repeat
```

## Summary

| Scenario | Approach |
|----------|----------|
| Pattern is consistent across rows | Derive rules → single SQL UPDATE |
| Multiple conditions / tiers | Multiple targeted UPDATEs |
| Rules unclear | Sample + aggregate → propose rules → confirm |
| Truly row-by-row judgment needed | Batch 50–100 rows at a time |

**Default to rules. Batch only as a last resort.**
