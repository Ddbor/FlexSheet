import {
  computeTableFormatCellStyle,
  normalizeSelectionRange,
  TABLE_ACCENT_PALETTES,
  type ParsedTableStyleCommand,
  type SelectionRange,
} from "@flexsheet/core";

function argbToCssRgb(argb: string): string {
  const u = argb.trim().toUpperCase();
  if (u.length === 8 && /^FF[0-9A-F]{6}$/.test(u)) {
    const r = parseInt(u.slice(2, 4), 16);
    const g = parseInt(u.slice(4, 6), 16);
    const b = parseInt(u.slice(6, 8), 16);
    return `rgb(${r},${g},${b})`;
  }
  return "#000000";
}

const THUMB_RANGE: SelectionRange = { startRow: 0, endRow: 5, startCol: 0, endCol: 4 };

/** 迷你表格预览（与 `computeTableFormatCellStyle` 一致，假定含标题行）。 */
export function createTableStyleThumbnailSvg(parsed: ParsedTableStyleCommand): SVGSVGElement {
  const palette = TABLE_ACCENT_PALETTES[parsed.col];
  const n = normalizeSelectionRange(THUMB_RANGE);
  const ns = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(ns, "svg");
  svg.setAttribute("viewBox", "0 0 50 60");
  svg.setAttribute("width", "50");
  svg.setAttribute("height", "60");
  svg.setAttribute("aria-hidden", "true");
  const borderCss = argbToCssRgb(palette.border);

  const cols = n.endCol - n.startCol + 1;
  const rows = n.endRow - n.startRow + 1;
  const cw = 50 / cols;
  const ch = 60 / rows;

  for (let r = n.startRow; r <= n.endRow; r++) {
    for (let c = n.startCol; c <= n.endCol; c++) {
      const st = computeTableFormatCellStyle(parsed, palette, n, true, r, c);
      const fill = st.fillArgb !== undefined ? argbToCssRgb(st.fillArgb) : "#ffffff";
      const rect = document.createElementNS(ns, "rect");
      rect.setAttribute("x", String((c - n.startCol) * cw));
      rect.setAttribute("y", String((r - n.startRow) * ch));
      rect.setAttribute("width", String(cw + 0.5));
      rect.setAttribute("height", String(ch + 0.5));
      rect.setAttribute("fill", fill);
      rect.setAttribute("stroke", borderCss);
      rect.setAttribute("stroke-width", "0.6");
      svg.appendChild(rect);
    }
  }

  return svg;
}

export function tableStyleCommandId(section: "light" | "medium" | "dark", row: number, col: number): string {
  return `home.style.table.${section}.r${row}c${col}`;
}
