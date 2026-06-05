import { Worksheet } from "exceljs";
import { CellValue, Sheet } from "./types.js";

export class MetadataSheet implements Sheet {
  constructor(
    public readonly label: string,
    private config: {
      /** @default 25 */
      keyWidth?: number;
      /** @default 25 */
      valueWidth?: number;
      entries: [CellValue, CellValue][];
    },
  ) {}

  async render(ws: Worksheet): Promise<void> {
    ws.columns = [{ width: this.config.keyWidth ?? 25 }, { width: this.config.valueWidth ?? 25 }];

    for (let i = 0; i < this.config.entries.length; i++) {
      const [key, value] = this.config.entries[i]!;
      const row = ws.getRow(i + 1);

      const keyCell = row.getCell(1);
      keyCell.value = key;

      const valueCell = row.getCell(2);
      valueCell.value = value;

      row.commit();
    }
  }
}
