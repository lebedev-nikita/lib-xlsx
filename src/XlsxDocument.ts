import exceljs from "exceljs";
import { Sheet } from "./types.js";
import { Buffer } from "node:buffer";

export class XlsxDocument {
  constructor(private sheets: Sheet[]) {}

  async renderToBuffer(): Promise<Buffer> {
    const wb = new exceljs.Workbook();
    await this.writeData(wb);
    return Buffer.from(await wb.xlsx.writeBuffer());
  }

  async renderToFile(name: string) {
    const wb = new exceljs.stream.xlsx.WorkbookWriter({ filename: name, useStyles: true });
    await this.writeData(wb);
    await wb.commit();
  }

  private async writeData(wb: exceljs.Workbook) {
    for (const sheet of this.sheets) {
      const ySplit = sheet.ySplit;
      const xSplit = sheet.xSplit;

      const ws =
        ySplit || xSplit
          ? wb.addWorksheet(sheet.label, { views: [{ state: "frozen", ySplit, xSplit }] })
          : wb.addWorksheet(sheet.label);

      await sheet.render(ws);
      if (typeof ws.commit == "function") {
        ws.commit();
      }
    }
  }
}
