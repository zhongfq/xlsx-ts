---
name: xlsx-agent
description: Inspect, modify, and validate `.xlsx` workbooks through the local `xlsx-ts` command line interface while preserving workbook structure and untouched XML parts. Use when Codex needs to edit Excel-based config files, batch-update cells, rename sheets, manage defined names, or verify roundtrip safety in this repository instead of hand-editing workbook XML.
---

# Xlsx Agent

Use the local CLI as the primary execution surface for workbook work.

## Use The Local CLI

Run commands from the repository root.

Prefer:

```bash
npm run cli -- <subcommand> ...
```

Use the built binary only if the package has already been built and linked:

```bash
xlsx-ts <subcommand> ...
```

## Follow This Workflow

1. Inspect the workbook before making changes.
2. Read back the exact sheet, cell, or workbook metadata that will be modified.
3. Use `set` for a single cell edit.
4. Use `apply` with an `ops.json` document for multi-step edits.
5. Validate the final workbook with `validate`.
6. Re-read important cells or workbook metadata after writing when the user cares about exact results.

## Inspect First

Start with:

```bash
npm run cli -- inspect path/to/file.xlsx
```

Use `get` when you need the current value, formula, style id, or number format of a single cell:

```bash
npm run cli -- get path/to/file.xlsx --sheet Sheet1 --cell B2
```

## Use Direct Commands For Small Edits

Write a plain string:

```bash
npm run cli -- set path/to/file.xlsx --sheet Sheet1 --cell B2 --text "hello" --output out.xlsx
```

Write a JSON scalar:

```bash
npm run cli -- set path/to/file.xlsx --sheet Sheet1 --cell B2 --value 123 --output out.xlsx
```

Write a formula with cached value:

```bash
npm run cli -- set path/to/file.xlsx --sheet Sheet1 --cell C2 --formula "B2*0.9" --cached-value 110.7 --output out.xlsx
```

Clear a cell:

```bash
npm run cli -- set path/to/file.xlsx --sheet Sheet1 --cell D2 --clear --output out.xlsx
```

Use `--in-place` only when the user clearly wants to overwrite the source workbook.

## Use `apply` For Multi-Step Edits

Create a JSON document, then apply it:

```bash
npm run cli -- apply path/to/file.xlsx --ops /tmp/xlsx-agent-ops.json --output out.xlsx
```

Prefer `apply` when the user asks for any of these patterns:

- Update several cells in one pass.
- Replace or append table-like records by header name.
- Rename sheets and then update workbook metadata.
- Add or delete sheets.
- Add or remove defined names.
- Combine value edits with active-sheet or visibility changes.

Read [ops-schema.md](./references/ops-schema.md) for the supported action shapes.

## Validate After Writing

Run:

```bash
npm run cli -- validate path/to/file.xlsx
```

Use validation whenever the user cares about style preservation, untouched workbook parts, or safe config updates.

## Use `table` For Structured Sheets

When a sheet has explicit header rows, metadata rows, and a later data section, prefer the `table` command group.

Examples:

```bash
npm run cli -- table inspect path/to/file.xlsx --sheet main --header-row 1 --data-start-row 6
npm run cli -- table list path/to/file.xlsx --sheet main --header-row 1 --data-start-row 6
npm run cli -- table get path/to/file.xlsx --sheet main --header-row 1 --data-start-row 6 --key-field id --key 1001
npm run cli -- table upsert path/to/file.xlsx --sheet main --header-row 1 --data-start-row 6 --key-field id --record '{"id":1001,"desc":"..."}' --in-place
npm run cli -- table sync path/to/file.xlsx --sheet main --header-row 1 --data-start-row 6 --key-field id --from-json main.json --in-place
```

Use this command group when rows between the header row and data rows must be preserved as-is.

## Use The Record Commands For Config Tables

When a sheet behaves like a config table with headers in row 1, prefer the record commands over manual cell-by-cell edits.

Prefer the high-level `config-table` command group when the task is really about config data rather than generic worksheet editing.

Examples:

```bash
npm run cli -- config-table sync path/to/file.xlsx --sheet Config --from-json config.json --output out.xlsx
npm run cli -- config-table init path/to/file.xlsx --sheet Config --headers '["Key","Value"]' --output out.xlsx
npm run cli -- config-table upsert out.xlsx --sheet Config --field Key --record '{"Key":"timeout","Value":"30"}' --in-place
npm run cli -- config-table get out.xlsx --sheet Config --field Key --text timeout
npm run cli -- config-table delete out.xlsx --sheet Config --field Key --text timeout --in-place
npm run cli -- config-table list out.xlsx --sheet Config
```

Use `config-table sync` when the user already has JSON config data and wants the workbook updated from that source in one pass.

Use the lower-level record commands only when the high-level command group is not expressive enough.

Examples:

```bash
npm run cli -- set-headers path/to/file.xlsx --sheet Config --headers '["Key","Value"]' --output out.xlsx
npm run cli -- add-record out.xlsx --sheet Config --record '{"Key":"timeout","Value":"30"}' --in-place
npm run cli -- set-records out.xlsx --sheet Config --records '[{"Key":"timeout","Value":"60"}]' --in-place
npm run cli -- records out.xlsx --sheet Config
```

Use `add-sheet` when the user explicitly wants a new worksheet:

```bash
npm run cli -- add-sheet path/to/file.xlsx --sheet Config --output out.xlsx
```

Use the direct worksheet commands when the user explicitly asks to rename or remove sheets:

```bash
npm run cli -- rename-sheet path/to/file.xlsx --from Sheet1 --to Config --output out.xlsx
npm run cli -- delete-sheet out.xlsx --sheet Scratch --in-place
```

## Use Style Commands For Common Formatting Work

Prefer the direct style commands for the common formatting tasks already supported by the CLI.

Examples:

```bash
npm run cli -- set-background-color path/to/file.xlsx --sheet Config --cell B2 --color FFFF0000 --output out.xlsx
npm run cli -- set-number-format out.xlsx --sheet Config --cell B2 --format '0.00%' --in-place
npm run cli -- copy-style out.xlsx --sheet Config --from B2 --to C2 --in-place
```

Prefer `apply` when formatting must be combined with several workbook edits in one transaction.

## Stay Inside The Supported Surface

Prefer the CLI over ad hoc TypeScript snippets or direct XML edits.

If the requested change is not supported by the current CLI:

1. Confirm that `inspect`, `get`, `set`, and `apply` cannot express it.
2. Extend the CLI in this repository.
3. Re-run the task through the CLI once the new command or action exists.

Do not patch workbook XML directly unless the user explicitly asks for a low-level intervention.
