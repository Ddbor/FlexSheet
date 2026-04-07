import { decodeTextFileBytes } from "@flexsheet/import-export";
import { describe, expect, it } from "vitest";

describe("decodeTextFileBytes", () => {
  it("decodes UTF-8 with BOM", () => {
    const enc = new TextEncoder();
    const raw = new Uint8Array([0xef, 0xbb, 0xbf, ...enc.encode("hello")]);
    expect(decodeTextFileBytes(raw)).toBe("hello");
  });

  it("decodes UTF-16 LE with BOM", () => {
    const body = "中文";
    const raw = new Uint8Array(2 + body.length * 2);
    raw[0] = 0xff;
    raw[1] = 0xfe;
    let o = 2;
    for (let i = 0; i < body.length; i++) {
      const code = body.charCodeAt(i);
      raw[o++] = code & 0xff;
      raw[o++] = (code >>> 8) & 0xff;
    }
    expect(decodeTextFileBytes(raw)).toBe("中文");
  });
});
