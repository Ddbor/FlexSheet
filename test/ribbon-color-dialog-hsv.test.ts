import { hsvToRgb, rgbToHsv } from "@flexsheet/toolbar";
import { describe, expect, it } from "vitest";

describe("ribbon-color-dialog hsv", () => {
  it("round-trips RGB through HSV", () => {
    const samples = [
      [157, 180, 219],
      [255, 0, 0],
      [0, 255, 0],
      [0, 0, 255],
      [128, 128, 128],
      [0, 0, 0],
      [255, 255, 255],
    ] as const;
    for (const [r, g, b] of samples) {
      const { h, s, v } = rgbToHsv(r, g, b);
      const back = hsvToRgb(h, s, v);
      expect(back.r).toBe(r);
      expect(back.g).toBe(g);
      expect(back.b).toBe(b);
    }
  });
});
