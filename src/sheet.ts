import type { CellValue } from "./types.js";
import type { Workbook } from "./workbook.js";

export class Sheet {
  readonly name: string;
  readonly path: string;
  readonly relationshipId: string;

  private readonly workbook: Workbook;

  constructor(
    workbook: Workbook,
    options: {
      name: string;
      path: string;
      relationshipId: string;
    },
  ) {
    this.workbook = workbook;
    this.name = options.name;
    this.path = options.path;
    this.relationshipId = options.relationshipId;
  }

  getCell(address: string): CellValue {
    return this.workbook.readCell(this, address);
  }

  setCell(address: string, value: CellValue): void {
    this.workbook.writeCell(this, address, value);
  }
}
