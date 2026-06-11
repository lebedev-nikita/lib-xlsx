import { assert, describe, expect, test } from "vitest";
import { colgroup, column } from "../src/helpers.js";
import { Group } from "../src/DataSheet.js";

describe("colgroup", () => {
  test("2 params", () => {
    const group = colgroup("abc", [column("col_id", "col_label")]);
    assert(group instanceof Group);

    expect(group.children.length).toBe(1);
    expect(group.columns.length).toBe(1);
    expect(group.columns.map((c) => [c.id, c.label])).toEqual([["col_id", "col_label"]]);
    expect(group.rowSpan).toBe(1);
  });

  test("3 params", () => {
    const group = colgroup("abc", { rowSpan: 3 }, []);
    assert(group instanceof Group);
    expect(group.rowSpan).toBe(3);
  });
});
