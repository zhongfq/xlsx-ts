import type { Sheet } from "./sheet.js";
import type { CellSnapshot, CellType, CellValue, SetFormulaOptions } from "./types.js";

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
