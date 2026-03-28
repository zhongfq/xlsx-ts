import type { Sheet } from "./sheet.js";
import type {
  CellBorderDefinition,
  CellBorderPatch,
  CellFillDefinition,
  CellFillPatch,
  CellFontDefinition,
  CellFontPatch,
  CellNumberFormatDefinition,
  CellSnapshot,
  CellStyleDefinition,
  CellStylePatch,
  CellType,
  CellValue,
  SetFormulaOptions,
} from "./types.js";

export class Cell {
  readonly address: string;

  private cachedRevision = -1;
  private cachedSnapshot?: CellSnapshot;
  private readonly sheet: Sheet;

  constructor(sheet: Sheet, address: string) {
    this.sheet = sheet;
    this.address = address;
  }

  get exists(): boolean {
    return this.getSnapshot().exists;
  }

  get formula(): string | null {
    return this.getSnapshot().formula;
  }

  get rawType(): string | null {
    return this.getSnapshot().rawType;
  }

  get styleId(): number | null {
    return this.getSnapshot().styleId;
  }

  get style(): CellStyleDefinition | null {
    return this.sheet.getStyle(this.address);
  }

  get font(): CellFontDefinition | null {
    return this.sheet.getFont(this.address);
  }

  get fill(): CellFillDefinition | null {
    return this.sheet.getFill(this.address);
  }

  get border(): CellBorderDefinition | null {
    return this.sheet.getBorder(this.address);
  }

  get numberFormat(): CellNumberFormatDefinition | null {
    return this.sheet.getNumberFormat(this.address);
  }

  get type(): CellType {
    return this.getSnapshot().type;
  }

  get value(): CellValue {
    return this.getSnapshot().value;
  }

  setFormula(formula: string, options: SetFormulaOptions = {}): void {
    this.sheet.setFormula(this.address, formula, options);
  }

  setValue(value: CellValue): void {
    this.sheet.setCell(this.address, value);
  }

  setStyleId(styleId: number | null): void {
    this.sheet.setStyleId(this.address, styleId);
  }

  setStyle(patch: CellStylePatch): number {
    return this.sheet.setStyle(this.address, patch);
  }

  setFont(patch: CellFontPatch): number {
    return this.sheet.setFont(this.address, patch);
  }

  setFill(patch: CellFillPatch): number {
    return this.sheet.setFill(this.address, patch);
  }

  setBorder(patch: CellBorderPatch): number {
    return this.sheet.setBorder(this.address, patch);
  }

  setNumberFormat(formatCode: string): number {
    return this.sheet.setNumberFormat(this.address, formatCode);
  }

  cloneStyle(patch: CellStylePatch = {}): number {
    return this.sheet.cloneStyle(this.address, patch);
  }

  private getSnapshot(): CellSnapshot {
    const revision = this.sheet.getRevision();
    if (this.cachedSnapshot && this.cachedRevision === revision) {
      return this.cachedSnapshot;
    }

    this.cachedSnapshot = this.sheet.readCellSnapshot(this.address);
    this.cachedRevision = revision;
    return this.cachedSnapshot;
  }
}
