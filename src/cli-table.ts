import { readFile } from "node:fs/promises";

import {
  assertCellRecord,
  assertCellRecordArray,
  assertPositiveInteger,
  assertRecord,
  assertString,
  assertStringArray,
  parseJsonDocument,
} from "./cli-json.js";
import type { CellRecord } from "./cli-json.js";
import type { CellValue } from "./types.js";
import type { Workbook } from "./workbook.js";

export type CliSheet = ReturnType<Workbook["getSheet"]>;

export interface ConfigTableRow {
  record: CellRecord;
  row: number;
}

export interface StructuredTableRow {
  record: CellRecord;
  row: number;
}

export interface TableProfile {
  dataStartRow: number;
  headerRow: number;
  keyFields?: string[];
  sheet: string;
}

export async function readConfigTableSyncInput(
  filePath: string,
  field: string,
  valueField: string,
): Promise<{
  headers?: string[];
  records: CellRecord[];
}> {
  const parsed = parseJsonDocument(await readFile(filePath, "utf8"), filePath);

  if (Array.isArray(parsed)) {
    return { records: assertCellRecordArray(parsed, `${filePath}`) };
  }

  const record = assertRecord(parsed, filePath);
  if (record.records !== undefined) {
    return {
      headers: record.headers === undefined ? undefined : assertStringArray(record.headers, `${filePath}.headers`),
      records: assertCellRecordArray(record.records, `${filePath}.records`),
    };
  }

  return {
    records: normalizeConfigObjectToRecords(record, filePath, field, valueField),
  };
}

export async function readTableProfiles(filePath: string): Promise<Record<string, TableProfile>> {
  const parsed = parseJsonDocument(await readFile(filePath, "utf8"), filePath);
  const root = assertRecord(parsed, filePath);
  const profiles = assertRecord(root.profiles, `${filePath}.profiles`);
  const next: Record<string, TableProfile> = {};

  for (const [name, rawProfile] of Object.entries(profiles)) {
    const profile = assertRecord(rawProfile, `${filePath}.profiles.${name}`);
    next[name] = {
      dataStartRow: assertPositiveInteger(profile.dataStartRow, `${filePath}.profiles.${name}.dataStartRow`),
      headerRow: assertPositiveInteger(profile.headerRow, `${filePath}.profiles.${name}.headerRow`),
      keyFields:
        profile.keyFields === undefined
          ? undefined
          : assertStringArray(profile.keyFields, `${filePath}.profiles.${name}.keyFields`),
      sheet: assertString(profile.sheet, `${filePath}.profiles.${name}.sheet`),
    };
  }

  return next;
}

export function getTableHeaders(sheet: CliSheet, headerRow: number): string[] {
  return sheet.getRow(headerRow).map((value) => (typeof value === "string" ? value : ""));
}

export function inferTableProfile(sheet: CliSheet): TableProfile {
  const headerRow = inferTableHeaderRow(sheet);
  const dataStartRow = inferTableDataStartRow(sheet, headerRow);
  const keyFields = inferTableKeyFields(getTableHeaders(sheet, headerRow));

  return {
    dataStartRow,
    headerRow,
    keyFields: keyFields.length > 0 ? keyFields : undefined,
    sheet: sheet.name,
  };
}

export function resolveTableKeyFields(
  sheet: CliSheet,
  headerRow: number,
  explicitKeyFields: string[],
): string[] {
  if (explicitKeyFields.length > 0) {
    return explicitKeyFields;
  }

  const inferred = inferTableKeyFields(getTableHeaders(sheet, headerRow));
  if (inferred.length > 0) {
    return inferred;
  }

  throw new Error("Unable to infer key fields; pass --key-field explicitly");
}

export function parseTableKey(source: string, keyFields: string[], label: string): CellRecord {
  const parsed = parseJsonDocument(source, label);

  if (keyFields.length === 1) {
    if (
      parsed === null ||
      typeof parsed === "string" ||
      typeof parsed === "number" ||
      typeof parsed === "boolean"
    ) {
      return { [keyFields[0]]: parsed };
    }
  }

  const record = assertCellRecord(parsed, label);
  return pickKeyRecord(record, keyFields);
}

export function pickKeyRecord(record: CellRecord, keyFields: string[]): CellRecord {
  const next: CellRecord = {};

  for (const keyField of keyFields) {
    if (!Object.hasOwn(record, keyField)) {
      throw new Error(`Record is missing key field: ${keyField}`);
    }

    next[keyField] = record[keyField] ?? null;
  }

  return next;
}

export function resolveConfigTableHeaders(
  sheet: CliSheet,
  headerRow: number,
  explicitHeaders: string[] | undefined,
  records: CellRecord[],
): string[] {
  if (explicitHeaders && explicitHeaders.length > 0) {
    return explicitHeaders;
  }

  const inferredHeaders = inferHeadersFromRecords(records);
  if (inferredHeaders.length > 0) {
    return inferredHeaders;
  }

  const existingHeaders = trimTrailingEmptyStrings(sheet.getHeaders(headerRow));
  if (existingHeaders.length > 0) {
    return existingHeaders;
  }

  throw new Error("Unable to determine headers; provide --headers or include records with keys");
}

export function normalizeConfigObjectToRecords(
  value: Record<string, unknown>,
  label: string,
  field: string,
  valueField: string,
): CellRecord[] {
  const records: CellRecord[] = [];

  for (const [key, entry] of Object.entries(value)) {
    if (
      entry === null ||
      typeof entry === "string" ||
      typeof entry === "number" ||
      typeof entry === "boolean"
    ) {
      records.push({
        [field]: key,
        [valueField]: entry,
      });
      continue;
    }

    const nested = assertCellRecord(entry, `${label}.${key}`);
    records.push({
      [field]: key,
      ...nested,
    });
  }

  return records;
}

export function listStructuredTableRows(
  sheet: CliSheet,
  headerRow: number,
  dataStartRow: number,
): StructuredTableRow[] {
  const rows: StructuredTableRow[] = [];

  for (let row = dataStartRow; row <= sheet.rowCount; row += 1) {
    const record = sheet.getRecord(row, headerRow);
    if (record !== null) {
      rows.push({ record, row });
    }
  }

  return rows;
}

export function findStructuredTableRow(
  sheet: CliSheet,
  headerRow: number,
  dataStartRow: number,
  keyFields: string[],
  key: CellRecord,
): StructuredTableRow | null {
  return (
    listStructuredTableRows(sheet, headerRow, dataStartRow).find((row) => matchesKey(row.record, keyFields, key)) ??
    null
  );
}

export function writeStructuredTableRecord(
  sheet: CliSheet,
  headerRow: number,
  rowNumber: number,
  record: CellRecord,
): void {
  const headers = getTableHeaders(sheet, headerRow);

  for (let columnIndex = 0; columnIndex < headers.length; columnIndex += 1) {
    const header = headers[columnIndex];
    if (header.length === 0) {
      continue;
    }

    const nextValue = Object.hasOwn(record, header) ? record[header] ?? null : null;
    sheet.setCell(rowNumber, columnIndex + 1, nextValue);
  }
}

export function writeStructuredTableRecords(
  sheet: CliSheet,
  headerRow: number,
  dataStartRow: number,
  records: CellRecord[],
): void {
  const existingRows = listStructuredTableRows(sheet, headerRow, dataStartRow).map((row) => row.row);

  for (let index = 0; index < records.length; index += 1) {
    writeStructuredTableRecord(sheet, headerRow, dataStartRow + index, records[index]);
  }

  const keepRows = new Set(records.map((_, index) => dataStartRow + index));
  const rowsToDelete = existingRows.filter((row) => !keepRows.has(row));
  rowsToDelete.sort((left, right) => right - left);

  for (const row of rowsToDelete) {
    sheet.deleteRecord(row, headerRow);
  }
}

export function listConfigTableRows(
  sheet: CliSheet,
  headerRow: number,
): ConfigTableRow[] {
  const rows: ConfigTableRow[] = [];

  for (let row = headerRow + 1; row <= sheet.rowCount; row += 1) {
    const record = sheet.getRecord(row, headerRow);
    if (record !== null) {
      rows.push({ record, row });
    }
  }

  return rows;
}

export function findConfigTableRow(
  sheet: CliSheet,
  headerRow: number,
  field: string,
  value: CellValue,
): ConfigTableRow | null {
  return listConfigTableRows(sheet, headerRow).find((row) => row.record[field] === value) ?? null;
}

export function inferProfileName(filePath: string, sheetName: string): string {
  const normalized = filePath.replaceAll("\\", "/");
  const fileName = normalized.slice(normalized.lastIndexOf("/") + 1);
  const withoutExtension = fileName.replace(/\.[^.]+$/, "");
  return `${withoutExtension}#${sheetName}`;
}

function inferHeadersFromRecords(records: CellRecord[]): string[] {
  const headers: string[] = [];
  const seen = new Set<string>();

  for (const record of records) {
    for (const key of Object.keys(record)) {
      if (!seen.has(key)) {
        seen.add(key);
        headers.push(key);
      }
    }
  }

  return headers;
}

function inferTableHeaderRow(sheet: CliSheet): number {
  const maxRow = Math.min(sheet.rowCount, 20);
  let bestRow = 0;
  let bestScore = Number.NEGATIVE_INFINITY;

  for (let row = 1; row <= maxRow; row += 1) {
    const score = scoreHeaderRowCandidate(sheet, row);
    if (score > bestScore) {
      bestScore = score;
      bestRow = row;
    }
  }

  if (bestRow === 0 || bestScore < 4) {
    throw new Error(`Unable to infer table header row for sheet: ${sheet.name}`);
  }

  return bestRow;
}

function inferTableDataStartRow(sheet: CliSheet, headerRow: number): number {
  for (let row = headerRow + 1; row <= sheet.rowCount; row += 1) {
    const values = sheet.getRow(row);
    if (isRowEmpty(values)) {
      continue;
    }

    const firstValue = values[0];
    if (typeof firstValue === "string" && (isMetadataMarker(firstValue) || looksLikeTypeDescriptor(firstValue))) {
      continue;
    }

    if (sheet.getRecord(row, headerRow) !== null) {
      return row;
    }
  }

  throw new Error(`Unable to infer table data start row for sheet: ${sheet.name}`);
}

function scoreHeaderRowCandidate(sheet: CliSheet, row: number): number {
  const values = sheet.getRow(row);
  const headers = trimTrailingEmptyStrings(values.map((value) => (typeof value === "string" ? value.trim() : "")));
  const nonEmptyHeaders = headers.filter((header) => header.length > 0);

  if (nonEmptyHeaders.length < 2) {
    return Number.NEGATIVE_INFINITY;
  }

  if (nonEmptyHeaders.some((header) => isMetadataMarker(header))) {
    return Number.NEGATIVE_INFINITY;
  }

  const uniqueHeaders = new Set(nonEmptyHeaders);
  let score = nonEmptyHeaders.length * 2;

  if (uniqueHeaders.size === nonEmptyHeaders.length) {
    score += 3;
  }

  if (nonEmptyHeaders.some((header) => header === "id" || header === "key" || /^key\d+$/.test(header))) {
    score += 4;
  }

  if (headers[0]?.startsWith("@")) {
    score -= 8;
  }

  return score;
}

function inferTableKeyFields(headers: string[]): string[] {
  const trimmedHeaders = headers.map((header) => header.trim()).filter((header) => header.length > 0);
  const compositeKeys: string[] = [];

  for (let index = 1; ; index += 1) {
    const name = `key${index}`;
    if (!trimmedHeaders.includes(name)) {
      break;
    }

    compositeKeys.push(name);
  }

  if (compositeKeys.length > 0) {
    return compositeKeys;
  }

  if (trimmedHeaders.includes("key")) {
    return ["key"];
  }

  if (trimmedHeaders.includes("id")) {
    return ["id"];
  }

  return [];
}

function isRowEmpty(values: CellValue[]): boolean {
  return values.every((value) => value === null || value === "");
}

function isMetadataMarker(value: string): boolean {
  return value === "auto" || value === ">>" || value === "!!!" || value === "###";
}

function looksLikeTypeDescriptor(value: string): boolean {
  const normalized = value.trim();
  return (
    normalized === "int" ||
    normalized === "string" ||
    normalized === "bool" ||
    normalized === "float" ||
    normalized === "number" ||
    normalized === "table" ||
    normalized === "items" ||
    normalized === "json" ||
    normalized === "int?" ||
    normalized === "string?" ||
    normalized === "bool?" ||
    normalized === "float?" ||
    normalized === "number?" ||
    normalized === "table?" ||
    normalized === "items?" ||
    normalized === "json?" ||
    normalized === "int[]" ||
    normalized === "string[]" ||
    normalized === "bool[]" ||
    normalized === "float[]" ||
    normalized === "number[]" ||
    normalized === "table[]" ||
    normalized === "items[]" ||
    normalized === "json[]" ||
    /^@[a-zA-Z_][a-zA-Z0-9_]*$/.test(normalized)
  );
}

function matchesKey(record: CellRecord, keyFields: string[], key: CellRecord): boolean {
  return keyFields.every((field) => record[field] === key[field]);
}

function trimTrailingEmptyStrings(values: string[]): string[] {
  let end = values.length;

  while (end > 0 && values[end - 1] === "") {
    end -= 1;
  }

  return values.slice(0, end);
}
