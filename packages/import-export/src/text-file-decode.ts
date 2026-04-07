/**
 * 将本地文件字节解码为字符串：按 BOM 识别 UTF-16 LE / UTF-16 BE，否则按 UTF-8（含 UTF-8 BOM）。
 */
export function decodeTextFileBytes(bytes: Uint8Array): string {
  if (bytes.length >= 2 && bytes[0] === 0xff && bytes[1] === 0xfe) {
    return new TextDecoder("utf-16le", { fatal: false }).decode(bytes);
  }
  if (bytes.length >= 2 && bytes[0] === 0xfe && bytes[1] === 0xff) {
    return new TextDecoder("utf-16be", { fatal: false }).decode(bytes);
  }
  return new TextDecoder("utf-8", { fatal: false }).decode(bytes);
}
