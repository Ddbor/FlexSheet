import { describe, expect, it } from "vitest";
import { normalizeSelectionRange, selectionRangeContains, Worksheet } from "@flexsheet/core";
import { SelectionModel } from "@flexsheet/selection";

describe("normalizeSelectionRange", () => {
  it("returns inclusive bounds when start is before end", () => {
    expect(
      normalizeSelectionRange({
        startRow: 1,
        startCol: 2,
        endRow: 3,
        endCol: 4,
      }),
    ).toEqual({ startRow: 1, startCol: 2, endRow: 3, endCol: 4 });
  });

  it("swaps when end is before start", () => {
    expect(
      normalizeSelectionRange({
        startRow: 5,
        startCol: 7,
        endRow: 2,
        endCol: 1,
      }),
    ).toEqual({ startRow: 2, startCol: 1, endRow: 5, endCol: 7 });
  });
});

describe("selectionRangeContains", () => {
  it("returns true for interior and edges of normalized range", () => {
    const r = { startRow: 1, startCol: 1, endRow: 3, endCol: 3 };
    expect(selectionRangeContains(r, 2, 2)).toBe(true);
    expect(selectionRangeContains(r, 1, 1)).toBe(true);
    expect(selectionRangeContains(r, 3, 3)).toBe(true);
  });

  it("returns false outside range", () => {
    const r = { startRow: 0, startCol: 0, endRow: 1, endCol: 1 };
    expect(selectionRangeContains(r, -1, 0)).toBe(false);
    expect(selectionRangeContains(r, 2, 0)).toBe(false);
    expect(selectionRangeContains(r, 0, 2)).toBe(false);
  });
});

describe("SelectionModel", () => {
  it("selectCell sets single-cell selection and clamps to sheet", () => {
    const sheet = new Worksheet("S", 5, 4);
    const model = new SelectionModel(() => sheet);
    model.selectCell(10, 10);
    expect(model.getActiveCell()).toEqual({ row: 4, col: 3 });
    expect(model.getAnchor()).toEqual({ row: 4, col: 3 });
    expect(model.getNormalizedRange()).toEqual({
      startRow: 4,
      startCol: 3,
      endRow: 4,
      endCol: 3,
    });
  });

  it("extendFocusTo keeps anchor and updates focus", () => {
    const sheet = new Worksheet("S", 10, 10);
    const model = new SelectionModel(() => sheet);
    model.selectCell(2, 2);
    model.extendFocusTo(4, 5);
    expect(model.getAnchor()).toEqual({ row: 2, col: 2 });
    expect(model.getActiveCell()).toEqual({ row: 4, col: 5 });
    expect(model.getNormalizedRange()).toEqual({
      startRow: 2,
      startCol: 2,
      endRow: 4,
      endCol: 5,
    });
  });

  it("moveFocus without extend moves anchor and focus together", () => {
    const sheet = new Worksheet("S", 10, 10);
    const model = new SelectionModel(() => sheet);
    model.selectCell(5, 5);
    model.moveFocus(0, 1, false);
    expect(model.getAnchor()).toEqual({ row: 5, col: 6 });
    expect(model.getActiveCell()).toEqual({ row: 5, col: 6 });
  });

  it("moveFocus with extend only moves focus", () => {
    const sheet = new Worksheet("S", 10, 10);
    const model = new SelectionModel(() => sheet);
    model.selectCell(3, 3);
    model.extendFocusTo(5, 5);
    model.moveFocus(-1, 0, true);
    expect(model.getAnchor()).toEqual({ row: 3, col: 3 });
    expect(model.getActiveCell()).toEqual({ row: 4, col: 5 });
  });

  it("syncWithSheet clamps indices when dimensions shrink", () => {
    let sheet: Worksheet = new Worksheet("S", 20, 20);
    const model = new SelectionModel(() => sheet);
    model.selectCell(15, 15);
    sheet = new Worksheet("S", 5, 5);
    model.syncWithSheet();
    expect(model.getActiveCell()).toEqual({ row: 4, col: 4 });
  });

  it("no-ops when sheet getter returns undefined", () => {
    const model = new SelectionModel(() => undefined);
    model.selectCell(0, 0);
    expect(model.getActiveCell()).toEqual({ row: 0, col: 0 });
  });
});
