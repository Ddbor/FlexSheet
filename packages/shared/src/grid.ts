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
