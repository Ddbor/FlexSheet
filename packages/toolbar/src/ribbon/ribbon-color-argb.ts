/** `#rrggbb` 或 `#rgb` → 不透明 OOXML ARGB（`FF` + RRGGBB 大写）。 */
export function cssHexToFillArgb(hex: string): string {
  const t = hex.trim().replace(/^#/, "").toLowerCase();
  let r: string;
  let g: string;
  let b: string;
  if (t.length === 3) {
    r = t[0]! + t[0]!;
    g = t[1]! + t[1]!;
    b = t[2]! + t[2]!;
  } else if (t.length === 6) {
    r = t.slice(0, 2);
    g = t.slice(2, 4);
    b = t.slice(4, 6);
  } else {
    return "FF000000";
  }
  return `FF${r}${g}${b}`.toUpperCase();
}

/** 8 位 ARGB → 供 `<input type="color">` 与 CSS 色条使用的 `#rrggbb`（忽略 alpha，全透明时返回 `#ffffff`）。 */
export function argb8ToCssHex6(argb: string): string {
  const s = argb.trim();
  if (!/^[\dA-Fa-f]{8}$/.test(s)) {
    return "#000000";
  }
  const a = parseInt(s.slice(0, 2), 16);
  if (a === 0) {
    return "#ffffff";
  }
  return `#${s.slice(2, 8).toLowerCase()}`;
}

/** 8 位 ARGB → CSS `background`（支持半透明）。 */
export function argb8ToStripeCss(argb: string | undefined): string | null {
  if (argb === undefined || argb === "") {
    return null;
  }
  const s = argb.trim();
  if (!/^[\dA-Fa-f]{8}$/.test(s)) {
    return null;
  }
  const a = parseInt(s.slice(0, 2), 16);
  const r = parseInt(s.slice(2, 4), 16);
  const g = parseInt(s.slice(4, 6), 16);
  const b = parseInt(s.slice(6, 8), 16);
  if (a === 0) {
    return null;
  }
  if (a < 255) {
    return `rgba(${r},${g},${b},${a / 255})`;
  }
  return `#${s.slice(2, 8).toLowerCase()}`;
}
