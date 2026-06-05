import exceljs from "exceljs";
import { describe, expect, test } from "vitest";
import { DataSheet, Group } from "../src/DataSheet.js";
import { XlsxDocument } from "../src/XlsxDocument.js";
import { MetadataSheet } from "../src/MetadataSheet.js";
import { Buffer } from "node:buffer";
import { Readable } from "node:stream";
import { column } from "../src/helpers.js";

async function readWorkbook(input: string | Buffer) {
  const wb = new exceljs.Workbook();
  if (typeof input == "string") {
    await wb.xlsx.readFile(input);
  } else {
    await wb.xlsx.read(Readable.from(input));
  }
  return wb;
}

async function readWorksheet(input: string | Buffer, sheetname: string) {
  const wb = await readWorkbook(input);
  const ws = wb.getWorksheet(sheetname);
  if (!ws) throw new Error(`worksheet not found: "${sheetname}"`);
  return ws;
}

describe("DataSheet", () => {
  const WORKSHEET_NAME = "my datasheet";
  const ELEMS = [1, 2, 3, 4, 5, 6];

  test("basic usage", async () => {
    const sheet = new DataSheet(WORKSHEET_NAME, {
      elems: ELEMS,
      headers: [
        column("first", "Первая подгруппа"),
        column("second", "Вторая подгруппа"),
        column("third", "Третья подгруппа"),
      ],
      renderRow(elem, ctx) {
        ctx.setValue("first", elem);
        ctx.setValue("third", elem * 3);
        ctx.setValue("second", elem * 2);
      },
    });

    const doc = new XlsxDocument([sheet]);
    const buffer = await doc.renderToBuffer();

    const ws = await readWorksheet(buffer, WORKSHEET_NAME);

    expect(ws.getCell(1, 1).value).toBe("Первая подгруппа");
    expect(ws.getCell(1, 2).value).toBe("Вторая подгруппа");
    expect(ws.getCell(1, 3).value).toBe("Третья подгруппа");

    for (let i = 1; i < ELEMS.length; i++) {
      const elem = ELEMS[i]!;
      expect(ws.getCell(i + 2, 1).value).toBe(elem);
      expect(ws.getCell(i + 2, 2).value).toBe(elem * 2);
      expect(ws.getCell(i + 2, 3).value).toBe(elem * 3);
    }
  });

  test("skipped column", async () => {
    const sheet = new DataSheet(WORKSHEET_NAME, {
      elems: ELEMS,
      headers: [
        column("first", "Первая подгруппа"),
        column("second", "Вторая подгруппа", { skipped: true }),
        column("third", "Третья подгруппа"),
      ],
      renderRow(elem, ctx) {
        ctx.setValue("first", elem);
        ctx.setValue("third", elem * 3);
        ctx.setValue("second", elem * 2);
      },
    });

    const doc = new XlsxDocument([sheet]);
    const buffer = await doc.renderToBuffer();

    const ws = await readWorksheet(buffer, WORKSHEET_NAME);

    expect(ws.getCell(1, 1).value).toBe("Первая подгруппа");
    expect(ws.getCell(1, 2).value).toBe("Третья подгруппа");

    for (let i = 1; i < ELEMS.length; i++) {
      const elem = ELEMS[i]!;
      expect(ws.getCell(i + 2, 1).value).toBe(elem);
      expect(ws.getCell(i + 2, 2).value).toBe(elem * 3);
    }
  });

  test("skipped all columns", async () => {
    const sheet = new DataSheet(WORKSHEET_NAME, {
      elems: ELEMS,
      headers: [
        column("first", "Первая подгруппа", { skipped: true }),
        column("second", "Вторая подгруппа", { skipped: true }),
        column("third", "Третья подгруппа", { skipped: true }),
      ],
      renderRow(elem, ctx) {
        ctx.setValue("first", elem);
        ctx.setValue("third", elem * 3);
        ctx.setValue("second", elem * 2);
      },
    });

    const doc = new XlsxDocument([sheet]);
    const buffer = await doc.renderToBuffer();

    const ws = await readWorksheet(buffer, WORKSHEET_NAME);

    expect(ws.getCell(1, 1).value).toBe(null);
    expect(ws.getCell(1, 2).value).toBe(null);
    expect(ws.getCell(1, 3).value).toBe(null);

    for (let i = 1; i < ELEMS.length; i++) {
      expect(ws.getCell(i + 2, 1).value).toBe(null);
      expect(ws.getCell(i + 2, 2).value).toBe(null);
      expect(ws.getCell(i + 2, 3).value).toBe(null);
    }
  });

  test("parent group", async () => {
    const sheet = new DataSheet(WORKSHEET_NAME, {
      elems: ELEMS,
      headers: [
        new Group({
          label: "Родительская группа",
          children: [
            column("first", "Первая подгруппа"),
            column("second", "Вторая подгруппа"),
            column("third", "Третья подгруппа"),
          ],
        }),
      ],
      renderRow(elem, ctx) {
        ctx.setValue("first", elem);
        ctx.setValue("third", elem * 3);
        ctx.setValue("second", elem * 2);
      },
    });

    const doc = new XlsxDocument([sheet]);
    const buffer = await doc.renderToBuffer();

    const ws = await readWorksheet(buffer, WORKSHEET_NAME);

    expect(ws.getCell(1, 1).value).toBe("Родительская группа");
    expect(ws.getCell(1, 2).value).toBe("Родительская группа");
    expect(ws.getCell(1, 3).value).toBe("Родительская группа");

    expect(ws.getCell(2, 1).value).toBe("Первая подгруппа");
    expect(ws.getCell(2, 2).value).toBe("Вторая подгруппа");
    expect(ws.getCell(2, 3).value).toBe("Третья подгруппа");

    for (let i = 1; i < ELEMS.length; i++) {
      const elem = ELEMS[i]!;
      expect(ws.getCell(i + 3, 1).value).toBe(elem);
      expect(ws.getCell(i + 3, 2).value).toBe(elem * 2);
      expect(ws.getCell(i + 3, 3).value).toBe(elem * 3);
    }
  });

  test("parent group with on child skipped", async () => {
    const sheet = new DataSheet(WORKSHEET_NAME, {
      elems: ELEMS,
      headers: [
        new Group({
          label: "Родительская группа",
          children: [
            column("first", "Первая подгруппа"),
            column("second", "Вторая подгруппа", { skipped: true }),
            column("third", "Третья подгруппа"),
          ],
        }),
      ],
      renderRow(elem, ctx) {
        ctx.setValue("first", elem);
        ctx.setValue("third", elem * 3);
        ctx.setValue("second", elem * 2);
      },
    });

    const doc = new XlsxDocument([sheet]);
    const buffer = await doc.renderToBuffer();

    const ws = await readWorksheet(buffer, WORKSHEET_NAME);

    expect(ws.getCell(1, 1).value).toBe("Родительская группа");
    expect(ws.getCell(1, 2).value).toBe("Родительская группа");
    expect(ws.getCell(1, 3).value).toBe(null);

    expect(ws.getCell(2, 1).value).toBe("Первая подгруппа");
    expect(ws.getCell(2, 2).value).toBe("Третья подгруппа");
    expect(ws.getCell(2, 3).value).toBe(null);

    for (let i = 1; i < ELEMS.length; i++) {
      const elem = ELEMS[i]!;
      expect(ws.getCell(i + 3, 1).value).toBe(elem);
      expect(ws.getCell(i + 3, 2).value).toBe(elem * 3);
      expect(ws.getCell(i + 3, 3).value).toBe(null);
    }
  });

  test("parent group with all children skipped", async () => {
    const sheet = new DataSheet(WORKSHEET_NAME, {
      elems: ELEMS,
      headers: [
        new Group({
          label: "Родительская группа",
          children: [
            column("first", "Первая подгруппа", { skipped: true }),
            column("second", "Вторая подгруппа", { skipped: true }),
            column("third", "Третья подгруппа", { skipped: true }),
          ],
        }),
      ],
      renderRow(elem, ctx) {
        ctx.setValue("first", elem);
        ctx.setValue("third", elem * 3);
        ctx.setValue("second", elem * 2);
      },
    });

    const doc = new XlsxDocument([sheet]);
    const buffer = await doc.renderToBuffer();

    const ws = await readWorksheet(buffer, WORKSHEET_NAME);

    expect(ws.getCell(1, 1).value).toBe(null);
    expect(ws.getCell(1, 2).value).toBe(null);
    expect(ws.getCell(1, 3).value).toBe(null);

    expect(ws.getCell(2, 1).value).toBe(null);
    expect(ws.getCell(2, 2).value).toBe(null);
    expect(ws.getCell(2, 3).value).toBe(null);

    for (let i = 1; i < ELEMS.length; i++) {
      expect(ws.getCell(i + 3, 1).value).toBe(null);
      expect(ws.getCell(i + 3, 2).value).toBe(null);
      expect(ws.getCell(i + 3, 3).value).toBe(null);
    }
  });

  test("parent group with split view", async () => {
    const sheet = new DataSheet(WORKSHEET_NAME, {
      elems: ELEMS,
      headers: [
        new Group({
          label: "Родительская группа",
          children: [
            column("first", "Первая подгруппа"),
            column("second", "Вторая подгруппа"),
            column("third", "Третья подгруппа"),
          ],
        }),
      ],
      freezeAfter: "second",
      renderRow(elem, ctx) {
        ctx.setValue("first", elem);
        ctx.setValue("third", elem * 3);
        ctx.setValue("second", elem * 2);
      },
    });

    const doc = new XlsxDocument([sheet]);
    const buffer = await doc.renderToBuffer();

    const ws = await readWorksheet(buffer, WORKSHEET_NAME);
    expect(ws.views[0]).toEqual(expect.objectContaining({ state: "frozen", xSplit: 2, ySplit: 2 }));
  });

  test("parent group with rowSpan=2", async () => {
    const sheet = new DataSheet(WORKSHEET_NAME, {
      elems: ELEMS,
      headers: [
        new Group({
          label: "Родительская группа",
          rowSpan: 2,
          children: [
            column("first", "Первая подгруппа"),
            column("second", "Вторая подгруппа"),
            column("third", "Третья подгруппа"),
          ],
        }),
      ],
      renderRow(elem, ctx) {
        ctx.setValue("first", elem);
        ctx.setValue("third", elem * 3);
        ctx.setValue("second", elem * 2);
      },
    });

    const doc = new XlsxDocument([sheet]);
    const buffer = await doc.renderToBuffer();

    const ws = await readWorksheet(buffer, WORKSHEET_NAME);

    expect(ws.getCell(1, 1).value).toBe("Родительская группа");
    expect(ws.getCell(1, 2).value).toBe("Родительская группа");
    expect(ws.getCell(1, 3).value).toBe("Родительская группа");
    expect(ws.getCell(2, 1).value).toBe("Родительская группа");
    expect(ws.getCell(2, 2).value).toBe("Родительская группа");
    expect(ws.getCell(2, 3).value).toBe("Родительская группа");

    expect(ws.getCell(3, 1).value).toBe("Первая подгруппа");
    expect(ws.getCell(3, 2).value).toBe("Вторая подгруппа");
    expect(ws.getCell(3, 3).value).toBe("Третья подгруппа");

    for (let i = 1; i < ELEMS.length; i++) {
      const elem = ELEMS[i]!;
      expect(ws.getCell(i + 4, 1).value).toBe(elem);
      expect(ws.getCell(i + 4, 2).value).toBe(elem * 2);
      expect(ws.getCell(i + 4, 3).value).toBe(elem * 3);
    }
  });

  test("column with rowSpan=2", async () => {
    const sheet = new DataSheet(WORKSHEET_NAME, {
      elems: ELEMS,
      headers: [
        new Group({
          label: "Родительская группа",
          children: [column("first", "Первая подгруппа"), column("second", "Вторая подгруппа")],
        }),
        column("third", "Третья подгруппа", { rowSpan: 2 }),
      ],
      renderRow(elem, ctx) {
        ctx.setValue("first", elem);
        ctx.setValue("third", elem * 3);
        ctx.setValue("second", elem * 2);
      },
    });

    const doc = new XlsxDocument([sheet]);
    const buffer = await doc.renderToBuffer();

    const ws = await readWorksheet(buffer, WORKSHEET_NAME);

    expect(ws.getCell(1, 1).value).toBe("Родительская группа");
    expect(ws.getCell(1, 2).value).toBe("Родительская группа");

    expect(ws.getCell(1, 3).value).toBe("Третья подгруппа");
    expect(ws.getCell(2, 3).value).toBe("Третья подгруппа");

    expect(ws.getCell(2, 1).value).toBe("Первая подгруппа");
    expect(ws.getCell(2, 2).value).toBe("Вторая подгруппа");

    for (let i = 1; i < ELEMS.length; i++) {
      const elem = ELEMS[i]!;
      expect(ws.getCell(i + 3, 1).value).toBe(elem);
      expect(ws.getCell(i + 3, 2).value).toBe(elem * 2);
      expect(ws.getCell(i + 3, 3).value).toBe(elem * 3);
    }
  });

  test("header styling", async () => {
    const sheet = new DataSheet(WORKSHEET_NAME, {
      elems: ELEMS,
      headers: [
        new Group({
          label: "Родительская группа",
          children: [column("first", "Первая подгруппа"), column("second", "Вторая подгруппа")],
        }),
        column("third", "Третья подгруппа", { rowSpan: 2, width: 40 }),
      ],
      renderRow(elem, ctx) {
        ctx.setValue("first", elem);
        ctx.setValue("third", elem * 3);
        ctx.setValue("second", elem * 2);
      },
    });

    const doc = new XlsxDocument([sheet]);
    const buffer = await doc.renderToBuffer();

    const ws = await readWorksheet(buffer, WORKSHEET_NAME);

    expect(ws.getColumn(1).width).toBe(25);
    expect(ws.getColumn(2).width).toBe(25);
    expect(ws.getColumn(3).width).toBe(40);

    expect(ws.getCell(1, 1).style.font?.name).toBe("Arial Black");
    expect(ws.getCell(1, 2).style.font?.name).toBe("Arial Black");

    expect(ws.getCell(1, 3).style.font?.name).toBe("Arial Black");
    expect(ws.getCell(2, 3).style.font?.name).toBe("Arial Black");

    expect(ws.getCell(2, 1).style.font?.name).toBe("Arial Black");
    expect(ws.getCell(2, 2).style.font?.name).toBe("Arial Black");
  });

  test("column.id must be unique", () => {
    expect(() => {
      const sheet = new DataSheet(WORKSHEET_NAME, {
        elems: ELEMS,
        headers: [column("same-id", "Первая подгруппа"), column("same-id", "Вторая подгруппа")],
        renderRow(elem, ctx) {},
      });
    }).toThrow(`column id met twice: "same-id"`);
  });
});

describe("MetadataSheet", () => {
  const WORKSHEET_NAME = "my metadatasheet";

  test("basic usage", async () => {
    const sheet = new MetadataSheet(WORKSHEET_NAME, {
      entries: [
        ["value 1", 1],
        ["value 2", "2"],
        ["value 3", null],
        ["value 4", undefined],
      ],
    });
    const doc = new XlsxDocument([sheet]);
    await doc.renderToBuffer();
  });
});

describe("XlsxDocument", () => {
  const WORKSHEET_NAME = "my datasheet";
  const ELEMS = [1, 2, 3, 4, 5, 6];

  test("renderToBuffer", async () => {
    const sheet = new DataSheet(WORKSHEET_NAME, {
      elems: ELEMS,
      headers: [
        column("first", "Первая подгруппа"),
        column("second", "Вторая подгруппа"),
        column("third", "Третья подгруппа"),
      ],
      renderRow(elem, ctx) {
        ctx.setValue("first", elem);
        ctx.setValue("third", elem * 3);
        ctx.setValue("second", elem * 2);
      },
    });

    const doc = new XlsxDocument([sheet]);
    const buffer = await doc.renderToBuffer();

    const ws = await readWorksheet(buffer, WORKSHEET_NAME);

    expect(ws.getCell(1, 1).value).toBe("Первая подгруппа");
    expect(ws.getCell(1, 2).value).toBe("Вторая подгруппа");
    expect(ws.getCell(1, 3).value).toBe("Третья подгруппа");

    for (let i = 1; i < ELEMS.length; i++) {
      const elem = ELEMS[i]!;
      expect(ws.getCell(i + 2, 1).value).toBe(elem);
      expect(ws.getCell(i + 2, 2).value).toBe(elem * 2);
      expect(ws.getCell(i + 2, 3).value).toBe(elem * 3);
    }
  });
});
