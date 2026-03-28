# Ops Schema

Use this file when preparing an `ops.json` document for:

```bash
npm run cli -- apply path/to/file.xlsx --ops ops.json --output out.xlsx
```

## File Shape

Use either of these forms:

```json
[
  { "type": "setCell", "sheet": "Sheet1", "cell": "A1", "value": "hello" }
]
```

```json
{
  "output": "out.xlsx",
  "actions": [
    { "type": "setCell", "sheet": "Sheet1", "cell": "A1", "value": "hello" }
  ]
}
```

`output` is optional. A CLI `--output` flag overrides it.

## Supported Actions

### setCell

```json
{ "type": "setCell", "sheet": "Sheet1", "cell": "B2", "value": 123 }
```

`value` must be a JSON scalar: string, number, boolean, or `null`.

### setBackgroundColor

```json
{ "type": "setBackgroundColor", "sheet": "Config", "cell": "B2", "color": "FFFF0000" }
```

Use `null` to clear a solid background fill:

```json
{ "type": "setBackgroundColor", "sheet": "Config", "cell": "B2", "color": null }
```

### setFormula

```json
{
  "type": "setFormula",
  "sheet": "Sheet1",
  "cell": "C2",
  "formula": "B2*0.9",
  "cachedValue": 110.7
}
```

`cachedValue` is optional and must be a JSON scalar when present.

### setNumberFormat

```json
{ "type": "setNumberFormat", "sheet": "Config", "cell": "B2", "formatCode": "0.00%" }
```

### clearCell

```json
{ "type": "clearCell", "sheet": "Sheet1", "cell": "D2" }
```

### renameSheet

```json
{ "type": "renameSheet", "from": "Sheet1", "to": "Config" }
```

### addSheet

```json
{ "type": "addSheet", "sheet": "Summary" }
```

### copyStyle

```json
{ "type": "copyStyle", "sheet": "Config", "from": "B2", "to": "C2" }
```

### setHeaders

```json
{ "type": "setHeaders", "sheet": "Config", "headers": ["Key", "Value"] }
```

Optional fields:

- `headerRow`
- `startColumn`

### deleteSheet

```json
{ "type": "deleteSheet", "sheet": "OldSheet" }
```

### addRecord

```json
{
  "type": "addRecord",
  "sheet": "Config",
  "record": { "Key": "timeout", "Value": "30" }
}
```

Optional field:

- `headerRow`

### addRecords

```json
{
  "type": "addRecords",
  "sheet": "Config",
  "records": [
    { "Key": "timeout", "Value": "30" },
    { "Key": "region", "Value": "ap-south-1" }
  ]
}
```

Optional field:

- `headerRow`

### setRecord

```json
{
  "type": "setRecord",
  "sheet": "Config",
  "row": 2,
  "record": { "Key": "timeout", "Value": "60" }
}
```

Optional field:

- `headerRow`

### setRecords

```json
{
  "type": "setRecords",
  "sheet": "Config",
  "records": [
    { "Key": "timeout", "Value": "60" }
  ]
}
```

Optional field:

- `headerRow`

### deleteRecord

```json
{ "type": "deleteRecord", "sheet": "Config", "row": 2 }
```

Optional field:

- `headerRow`

### deleteRecords

```json
{ "type": "deleteRecords", "sheet": "Config", "rows": [2, 4, 7] }
```

Optional field:

- `headerRow`

### setActiveSheet

```json
{ "type": "setActiveSheet", "sheet": "Config" }
```

### setSheetVisibility

```json
{ "type": "setSheetVisibility", "sheet": "Config", "visibility": "hidden" }
```

Allowed visibility values:

- `visible`
- `hidden`
- `veryHidden`

### setDefinedName

Global:

```json
{ "type": "setDefinedName", "name": "ConfigRange", "value": "Config!$A$1:$B$20" }
```

Sheet-scoped:

```json
{ "type": "setDefinedName", "name": "LocalValue", "scope": "Config", "value": "$B$2" }
```

### deleteDefinedName

Global:

```json
{ "type": "deleteDefinedName", "name": "ConfigRange" }
```

Sheet-scoped:

```json
{ "type": "deleteDefinedName", "name": "LocalValue", "scope": "Config" }
```

## Recommended Pattern

For multi-step workbook edits:

1. Run `inspect`.
2. Build `ops.json`.
3. Run `apply`.
4. Run `validate`.
5. Re-run `get` on critical cells if the user needs exact confirmation.
