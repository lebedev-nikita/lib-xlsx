import exceljs from "exceljs";
import { XLSX_MAX_ROWS } from "./consts.js";
import { CellValue, Sheet, Style } from "./types.js";

export type Context<TCellId extends string> = {
  rowIdx: number;
  elemIdx: number;
  setValue(
    id: TCellId,
    value: CellValue,
    opt?: {
      style?: Style;
    },
  ): void;
};

export interface HeaderItem<TColumnId extends string> {
  readonly label: string;
  readonly rowSpan: number;
  readonly children: readonly HeaderItem<TColumnId>[];
  readonly skipped: boolean;
  readonly columns: Column<TColumnId>[];
}

export class Group<TColumnId extends string> implements HeaderItem<TColumnId> {
  readonly label: string;
  readonly children: readonly HeaderItem<TColumnId>[];
  readonly rowSpan: number;
  readonly skipped: boolean = false;

  constructor(config: {
    label: string;
    children: HeaderItem<TColumnId>[];
    /** @default 1 */
    rowSpan?: number;
  }) {
    this.label = config.label;
    this.children = config.children;
    this.rowSpan = config.rowSpan ?? 1;
    this.skipped = this.children.every((child) => child.skipped);
  }

  get columns(): Column<TColumnId>[] {
    return this.children.flatMap((group) => group.columns);
  }
}

export class Column<const TColumnId extends string> implements HeaderItem<TColumnId> {
  readonly id: TColumnId;
  readonly label: string;
  readonly rowSpan: number;
  readonly children: readonly HeaderItem<TColumnId>[] = [];
  readonly skipped: boolean;
  readonly width: number;

  constructor(config: {
    label: string;
    id: TColumnId;
    /** @default false */
    skipped?: boolean;
    /** @default 1 */
    rowSpan?: number;
    /** @default 25 */
    width?: number;
  }) {
    this.id = config.id;
    this.label = config.label;
    this.rowSpan = config.rowSpan ?? 1;
    this.skipped = config.skipped ?? false;
    this.width = config.width ?? 25;
  }

  get columns() {
    return [this];
  }
}

// TODO: Добавить MetadataSheet

export class DataSheet<TColumnId extends string, TElem extends unknown> implements Sheet {
  private rowIdx = 1;
  private colIndexes = new Map<string, number>();

  constructor(
    public readonly label: string,
    private config: {
      headers: HeaderItem<TColumnId>[];
      freezeAfter?: NoInfer<TColumnId>;
      elems: Iterable<TElem> | AsyncIterable<TElem>;
      renderRow: (elem: TElem, ctx: Context<NoInfer<TColumnId>>) => void | Promise<void>;
    },
  ) {
    {
      // validate column ids
      const metIds = new Set<string>();
      for (const col of this.columns) {
        if (metIds.has(col.id)) {
          throw new Error(`column id met twice: "${col.id}"`);
        }
        metIds.add(col.id);
      }
    }

    {
      // map column ids to indexes of not skipped columns
      let columnIdx = 1;
      for (const column of this.notSkippedColumns) {
        this.colIndexes.set(column.id, columnIdx++);
      }
    }
  }

  private validateColumnIds() {}

  private get columns() {
    return this.config.headers.flatMap((group) => group.columns);
  }

  private get notSkippedColumns() {
    return this.columns.filter((col) => !col.skipped);
  }

  async render(ws: exceljs.Worksheet) {
    await this.renderHeaders(ws);
    await this.renderRows(ws);
  }

  private async renderHeaders(ws: exceljs.Worksheet) {
    type Coord = { x: number; y: number };

    function renderItem(start: { item: HeaderItem<string> } & Coord): Coord | null {
      if (start.item.skipped) return null;

      const cell = ws.getCell(start.y, start.x);
      cell.value = start.item.label;
      cell.style.font = { name: "Arial Black" };
      cell.style.alignment = { vertical: "middle", horizontal: "center" };

      const labelEnd: Coord = { x: start.x, y: start.y + start.item.rowSpan - 1 };
      const groupEnd: Coord = { ...labelEnd };

      if (start.item instanceof Group) {
        let first = true;
        for (const group of start.item.children) {
          const result = renderItem({
            item: group,
            x: first ? start.x : groupEnd.x + 1,
            y: labelEnd.y + 1,
          });
          if (!result) continue;
          groupEnd.x = Math.max(groupEnd.x, result.x);
          groupEnd.y = Math.max(groupEnd.y, result.y);
          first = false;
        }
      } else if (start.item instanceof Column) {
        const col = ws.columns[start.x - 1];
        if (col) {
          col.width = start.item.width;
        }
      }

      ws.mergeCells({ left: start.x, top: start.y, bottom: labelEnd.y, right: groupEnd.x });
      return groupEnd;
    }

    const end: Coord = { x: 0, y: 0 };
    for (const group of this.config.headers) {
      const result = renderItem({ item: group, x: end.x + 1, y: 1 });
      if (!result) continue;
      end.x = Math.max(end.x, result.x);
      end.y = Math.max(end.y, result.y);
    }

    for (let i = 1; i < +end.y; i++) {
      const row = ws.getRow(i);
      row.commit();
    }
    this.rowIdx = end.y + 1;
  }

  private async renderRows(ws: exceljs.Worksheet) {
    let elemIdx = 0;
    for await (const elem of this.config.elems) {
      if (this.rowIdx > XLSX_MAX_ROWS) break;

      const parent = this;
      const row = ws.getRow(this.rowIdx++);

      await this.config.renderRow(elem, {
        rowIdx: parent.rowIdx,
        elemIdx: elemIdx++,
        setValue(id, value, opt) {
          const j = parent.colIndexes.get(id);
          if (!j) return;
          const cell = row.getCell(j);
          cell.value = value;
          if (opt?.style?.font) {
            cell.style.font = {
              underline: opt.style.font.underline,
              bold: opt.style.font.bold,
              italic: opt.style.font.italic,
              name: opt.style.font.name,
              get color() {
                if (!opt.style?.font?.color) return undefined;
                if (!/^[0-9A-F]{8}$/i.test(opt.style.font.color)) {
                  throw new Error(`color "${opt.style.font.color}" does not match ARGB (AABBCCDD)`);
                }
                return { argb: opt.style.font.color };
              },
            };
          }
        },
      });

      row.commit();
    }
  }

  get ySplit(): number {
    function calculateDepth(group: HeaderItem<string>): number {
      if (group.skipped) return 0;
      if (group instanceof Column) return group.rowSpan;
      let maxDepth = 0;
      for (const child of group.children) {
        maxDepth = Math.max(maxDepth, calculateDepth(child));
      }
      return group.rowSpan + maxDepth;
    }

    let headerHeight = 0;
    for (const group of this.config.headers) {
      headerHeight = Math.max(headerHeight, calculateDepth(group));
    }
    return headerHeight;
  }

  get xSplit(): number {
    let ret = 0;

    for (const col of this.columns) {
      if (!col.skipped) ret++;
      if (col.id == this.config.freezeAfter) break;
    }

    return ret;
  }
}
export { XLSX_MAX_ROWS };
