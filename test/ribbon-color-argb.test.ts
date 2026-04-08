import { argb8ToCssHex6, argb8ToStripeCss, cssHexToFillArgb } from "@flexsheet/toolbar";
import { describe, expect, it } from "vitest";

describe("ribbon-color-argb", () => {
  it("cssHexToFillArgb expands 6-digit hex", () => {
    expect(cssHexToFillArgb("#4472c4")).toBe("FF4472C4");
    expect(cssHexToFillArgb("ed7d31")).toBe("FFED7D31");
  });

  it("argb8ToCssHex6 drops alpha for color input", () => {
    expect(argb8ToCssHex6("FF00B050")).toBe("#00b050");
    expect(argb8ToCssHex6("00000000")).toBe("#ffffff");
  });

  it("argb8ToStripeCss respects alpha", () => {
    expect(argb8ToStripeCss("FF0000FF")).toBe("#0000ff");
    expect(argb8ToStripeCss("800000FF")).toMatch(/^rgba\(0,\s*0,\s*255,\s*0\.5/);
    expect(argb8ToStripeCss("00000000")).toBe(null);
  });
});
