/**
 * 列索引与 Excel 风格列标（A、B、…、Z、AA）互转。
 */

export function columnIndexToLabel(index: number): string {
  if (index < 0) {
    return "";
  }
  let label = "";
  let i = index;
  while (i >= 0) {
    label = String.fromCharCode((i % 26) + 65) + label;
    i = Math.floor(i / 26) - 1;
  }
  return label;
}

/** 解析 Excel 列标（A、Z、AA）为 0 基列索引；无法解析时返回 `null`。 */
export function columnLabelToIndex(label: string): number | null {
  const s = label.trim().toUpperCase().replace(/[^A-Z]/g, "");
  if (s === "" || s.length > 6) {
    return null;
  }
  let n = 0;
  for (let i = 0; i < s.length; i++) {
    const c = s.charCodeAt(i);
    if (c < 65 || c > 90) {
      return null;
    }
    n = n * 26 + (c - 64);
  }
  return n - 1;
}
