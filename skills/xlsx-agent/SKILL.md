---
name: xlsx-agent
description: Edit and validate `.xlsx` workbooks in this repository through the local `xlsx-ts` CLI. Use for config-table updates, structured-sheet edits, sheet management, style changes, and roundtrip-safe workbook modifications instead of touching workbook XML directly.
---

# Xlsx Agent

Use the local CLI from the repository root:

```bash
npm run cli -- <subcommand> ...
```

## Default Workflow

1. Inspect before writing.
2. Read the exact sheet, cell, or table row that will change.
3. Apply the smallest fitting command.
4. Validate after writing.
5. Re-read key results when exact output matters.

Start with:

```bash
npm run cli -- inspect path/to/file.xlsx
npm run cli -- get path/to/file.xlsx --sheet Sheet1 --cell B2
npm run cli -- validate path/to/file.xlsx
```

Use `--in-place` only when the user clearly wants to overwrite the source file.

## Command Choice

Use `set` for single-cell edits:

```bash
npm run cli -- set path/to/file.xlsx --sheet Sheet1 --cell B2 --text "hello" --output out.xlsx
```

Use `apply` for multi-step edits:

```bash
npm run cli -- apply path/to/file.xlsx --ops /tmp/xlsx-agent-ops.json --output out.xlsx
```

Use `config-table` for simple header-based config sheets:

```bash
npm run cli -- config-table list path/to/file.xlsx --sheet Config
npm run cli -- config-table upsert path/to/file.xlsx --sheet Config --field Key --record '{"Key":"timeout","Value":"30"}' --in-place
```

Use `table` for structured sheets with header rows, metadata rows, and later data rows:

```bash
npm run cli -- table inspect path/to/file.xlsx --sheet main --header-row 1 --data-start-row 6
npm run cli -- table upsert path/to/file.xlsx --sheet main --header-row 1 --data-start-row 6 --key-field id --record '{"id":1001,"desc":"..."}' --in-place
```

Treat rows such as `auto`, `>>`, `!!!`, `###`, and `-` as structure to preserve, not built-in business semantics.

## Profiles

If `table-profiles.json` already exists, prefer `--profile`:

```bash
npm run cli -- table list res/task.xlsx --profile 'task#main'
npm run cli -- table get res/task.xlsx --profile 'task#conf' --key '"GATE_SIEGE_TIME"'
npm run cli -- table get res/task.xlsx --profile 'task#define' --key '{"key1":"TASK_TYPE","key2":"MAIN"}'
```

If profiles do not exist yet, generate them first:

```bash
npm run cli -- table generate-profiles res/task.xlsx
npm run cli -- table generate-profiles res/task.xlsx res/monster.xlsx --sheet-filter '^(main|conf)$' --output table-profiles.json
```

Generated names use `文件名#表名`, for example `task#main`.

## Limits

Prefer the CLI over ad hoc scripts or direct XML edits.

If the current CLI cannot express the requested change:

1. Confirm that existing commands are insufficient.
2. Extend the CLI in this repository.
3. Re-run the workbook change through the CLI.
