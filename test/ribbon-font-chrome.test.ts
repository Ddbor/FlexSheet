import { describe, expect, it } from "vitest";
import { cellStyleToRibbonHomeFontChrome } from "@flexsheet/toolbar";

describe("cellStyleToRibbonHomeFontChrome", () => {
  it("uses defaults when style is null", () => {
    const s = cellStyleToRibbonHomeFontChrome(null);
    expect(s.fontLabel).toBe("微软雅黑");
    expect(s.sizeLabel).toBe("11");
    expect(s.boldPressed).toBe(false);
    expect(s.italicPressed).toBe(false);
    expect(s.underlinePressed).toBe(false);
    expect(s.doubleUnderlinePressed).toBe(false);
    expect(s.fillStripeCss).toBe("#000000");
    expect(s.fontStripeCss).toBe("#ffffff");
  });

  it("maps known ribbon font stack to list label", () => {
    const s = cellStyleToRibbonHomeFontChrome({
      fontFamily: '"Microsoft YaHei", "微软雅黑", "PingFang SC", sans-serif',
      fontSizePt: 12,
      bold: true,
      italic: true,
      underline: "double",
    });
    expect(s.fontLabel).toBe("微软雅黑");
    expect(s.sizeLabel).toBe("12");
    expect(s.boldPressed).toBe(true);
    expect(s.italicPressed).toBe(true);
    expect(s.underlinePressed).toBe(false);
    expect(s.doubleUnderlinePressed).toBe(true);
    expect(s.fillStripeCss).toBe("#000000");
    expect(s.fontStripeCss).toBe("#ffffff");
  });

  it("shows single underline pressed", () => {
    const s = cellStyleToRibbonHomeFontChrome({ underline: "single" });
    expect(s.underlinePressed).toBe(true);
    expect(s.doubleUnderlinePressed).toBe(false);
    expect(s.fillStripeCss).toBe("#000000");
    expect(s.fontStripeCss).toBe("#ffffff");
  });

  it("maps fill and font ARGB to color stripes", () => {
    const s = cellStyleToRibbonHomeFontChrome({
      fillArgb: "FFFFFF00",
      fgArgb: "FFFF0000",
    });
    expect(s.fillStripeCss).toBe("#ffff00");
    expect(s.fontStripeCss).toBe("#ff0000");
  });

  it("uses transparent grid stripe only for fully transparent fill ARGB", () => {
    const s = cellStyleToRibbonHomeFontChrome({ fillArgb: "00000000" });
    expect(s.fillStripeCss).toBe(null);
    expect(s.fontStripeCss).toBe("#ffffff");
  });
});
