# excel-sql-skill

A skill that lets Claude Code query and edit **live Excel files** using SQL.

Uses **xlwings** (Excel process control) + **pandasql** (in-memory SQL) — no intermediate database files.

## Why?

When working with large Excel datasets, loading the whole file into context is wasteful and slow. This skill lets Claude Code:
- **Read** data with SQL queries (`SELECT * FROM Monsters WHERE hp > 1000`)
- **Write** changes back to the open Excel file (`UPDATE Monsters SET damage = damage * 1.2`)
- Excel **auto-recalculates** formulas — you see results in real-time

## Install

### 1. Install the skill

```bash
npx skills add github:yulcat/excel-sql-skill
```

Or clone manually into your project's `skills/` folder.

### 2. Install Python dependencies

```bash
pip install -r requirements.txt
```

### 3. Ensure Excel is open

The target Excel file must be open in Excel before attaching.

## Platform Support

| Platform | Live Edit | Notes |
|----------|-----------|-------|
| Windows  | ✅ Full   | COM automation, real-time cell updates |
| macOS    | ✅ Works  | AppleScript bridge, may need manual save |
| Linux    | ❌ No     | No Excel desktop app |

## Usage

See [SKILL.md](SKILL.md) for detailed usage instructions and examples.

## Architecture

```
Excel (open) ←→ xlwings ←→ pandas DataFrame ←→ pandasql (SQL) ←→ Claude Code
                  ↑                                                      ↓
           writes cells                                           reads/analyzes
           directly                                               via SQL queries
```

## License

MIT
