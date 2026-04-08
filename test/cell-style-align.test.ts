import { applyCellStylePatch, type CellStyle } from "@flexsheet/core";
import { describe, expect, it } from "vitest";

describe("CellStyle alignment patch", () => {
  it("applies and clears hAlign / vAlign", () => {
    const base: CellStyle = { hAlign: "center", vAlign: "top" };
    expect(applyCellStylePatch(base, { hAlign: "right" })).toEqual({
      hAlign: "right",
      vAlign: "top",
    });
    expect(applyCellStylePatch(base, { hAlign: null })).toEqual({ vAlign: "top" });
    expect(applyCellStylePatch(base, { vAlign: null })).toEqual({ hAlign: "center" });
    expect(applyCellStylePatch(base, { hAlign: null, vAlign: null })).toBeNull();
  });

  it("applies indentLevel and wrapText", () => {
    expect(applyCellStylePatch(null, { indentLevel: 3, wrapText: true })).toEqual({
      indentLevel: 3,
      wrapText: true,
    });
    expect(applyCellStylePatch({ indentLevel: 2 }, { indentLevel: null })).toBeNull();
    expect(applyCellStylePatch({ wrapText: true }, { wrapText: null })).toBeNull();
  });
});
