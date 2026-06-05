import exceljs from "exceljs";

export interface Sheet {
  label: string;
  xSplit?: number;
  ySplit?: number;
  render(ws: exceljs.Worksheet): Promise<void>;
}

export type CellValue =
  | number
  | string
  | null
  | undefined
  | exceljs.CellRichTextValue
  | exceljs.CellHyperlinkValue
  | exceljs.CellFormulaValue
  | exceljs.CellSharedFormulaValue;

export type Style = {
  font?: {
    /** @example color: "AABBCCDD" */
    color?: string;
    underline?: boolean;
    bold?: boolean;
    italic?: boolean;
    name?: string;
  };
};
