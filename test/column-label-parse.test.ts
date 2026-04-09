import { columnIndexToLabel, columnLabelToIndex } from "@flexsheet/shared";
import { describe, expect, it } from "vitest";

describe("columnLabelToIndex", () => {
  it("parses A, Z, AA", () => {
    expect(columnLabelToIndex("A")).toBe(0);
    expect(columnLabelToIndex("z")).toBe(25);
    expect(columnLabelToIndex("AA")).toBe(26);
  });

  it("round-trips with columnIndexToLabel for small indices", () => {
    for (let i = 0; i < 300; i++) {
      const label = columnIndexToLabel(i);
      expect(columnLabelToIndex(label)).toBe(i);
    }
  });

  it("returns null for invalid", () => {
    expect(columnLabelToIndex("")).toBe(null);
    expect(columnLabelToIndex("1")).toBe(null);
  });
});
