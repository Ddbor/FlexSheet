import { describe, expect, it } from "vitest";
import { Worksheet } from "@flexsheet/core";
import { buildFrozenLayout, hitTestHeadingPointer } from "@flexsheet/renderer";

describe("hitTestHeadingPointer", () => {
  const headerW = 40;
  const headerH = 24;

  it("returns selectAllCorner inside top-left chrome", () => {
    const sheet = new Worksheet("S", 20, 8);
    const layout = buildFrozenLayout(sheet, headerW, headerH, 600, 400, 0, 0, 1);
    expect(hitTestHeadingPointer(10, 10, sheet, layout, 0, 0, 1)).toEqual({
      kind: "selectAllCorner",
    });
  });

  it("returns columnHeader for letter strip", () => {
    const sheet = new Worksheet("S", 20, 8);
    const layout = buildFrozenLayout(sheet, headerW, headerH, 600, 400, 0, 0, 1);
    const h = hitTestHeadingPointer(headerW + 40, headerH / 2, sheet, layout, 0, 0, 1);
    expect(h?.kind).toBe("columnHeader");
    if (h?.kind === "columnHeader") {
      expect(h.col).toBe(0);
    }
  });

  it("returns rowHeader for row-number strip", () => {
    const sheet = new Worksheet("S", 20, 8);
    const layout = buildFrozenLayout(sheet, headerW, headerH, 600, 400, 0, 0, 1);
    const h = hitTestHeadingPointer(headerW / 2, headerH + 48, sheet, layout, 0, 0, 1);
    expect(h?.kind).toBe("rowHeader");
    if (h?.kind === "rowHeader") {
      expect(h.row).toBe(2);
    }
  });

  it("returns null for body", () => {
    const sheet = new Worksheet("S", 20, 8);
    const layout = buildFrozenLayout(sheet, headerW, headerH, 600, 400, 0, 0, 1);
    expect(hitTestHeadingPointer(headerW + 100, headerH + 100, sheet, layout, 0, 0, 1)).toBeNull();
  });
});
