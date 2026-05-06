import { formatCellDisplayWithStyle, type Worksheet } from "@flexsheet/core";
import {
  buildCellCanvasFont,
  cellStyleLogicalFontSizeBasePx,
  scaledFontSizePx,
  wrapCellLines,
} from "./canvas/canvas-renderer-utils.js";
import { bodyColumnAutoFilterTextReservePx } from "./hit-test/grid-hit-test.js";

const PAD = 4;
/** 与拖拽调整行列尺寸下限（flex-sheet）一致。 */
const MIN_LINE_SIZE = 8;

function clampIndentLevel(raw: number | undefined): number {
  if (raw === undefined || !Number.isFinite(raw)) {
    return 0;
  }
  return Math.max(0, Math.min(255, Math.round(raw)));
}

let measureCanvas: HTMLCanvasElement | null = null;

function measureCtx(): CanvasRenderingContext2D {
  if (typeof document === "undefined") {
    throw new Error("computeColumnAutoWidth/computeRowAutoHeight require a browser document");
  }
  if (measureCanvas === null) {
    measureCanvas = document.createElement("canvas");
  }
  const ctx = measureCanvas.getContext("2d");
  if (ctx === null) {
    throw new Error("Canvas 2D context unavailable");
  }
  return ctx;
}

/**
 * 按该列所有可见单元格内容计算列宽（与当前绘制用的字体度量一致）。
 */
export function computeColumnAutoWidth(sheet: Worksheet, col: number, viewZoom: number): number {
  if (col < 0 || col >= sheet.colCount || sheet.isColHidden(col)) {
    return sheet.getColWidth(col);
  }
  const ctx = measureCtx();
  let maxNeed = 0;
  for (let r = 0; r < sheet.rowCount; r++) {
    if (sheet.isRowHidden(r)) {
      continue;
    }
    const cell = sheet.getCell(r, col);
    const text = formatCellDisplayWithStyle(cell.value, cell.style);
    if (text === "") {
      continue;
    }
    const st = cell.style;
    ctx.font = buildCellCanvasFont(st, viewZoom);
    const fontPx = scaledFontSizePx(cellStyleLogicalFontSizeBasePx(st), viewZoom);
    const indentLv = clampIndentLevel(st?.indentLevel);
    const indentUnit = Math.max(6, fontPx * 0.55);
    const indentPx = indentLv * indentUnit;
    const lines = text.split("\n");
    let maxLineW = 0;
    for (const line of lines) {
      maxLineW = Math.max(maxLineW, ctx.measureText(line).width);
    }
    const filterReserve =
      bodyColumnAutoFilterTextReservePx(sheet, r, col) + sheet.getPivotPageFilterDropdownReservePx(r, col);
    const need = Math.ceil(indentPx + maxLineW + PAD * 2 + filterReserve);
    maxNeed = Math.max(maxNeed, need);
  }
  return Math.max(MIN_LINE_SIZE, maxNeed > 0 ? maxNeed : sheet.defaultColWidth);
}

/**
 * 按该行所有可见单元格内容计算行高（自动换行时按当前列宽折行）。
 */
export function computeRowAutoHeight(sheet: Worksheet, row: number, viewZoom: number): number {
  if (row < 0 || row >= sheet.rowCount || sheet.isRowHidden(row)) {
    return sheet.getRowHeight(row);
  }
  const ctx = measureCtx();
  let maxH = 0;
  for (let c = 0; c < sheet.colCount; c++) {
    if (sheet.isColHidden(c)) {
      continue;
    }
    const cell = sheet.getCell(row, c);
    const text = formatCellDisplayWithStyle(cell.value, cell.style);
    if (text === "") {
      continue;
    }
    const st = cell.style;
    ctx.font = buildCellCanvasFont(st, viewZoom);
    const fontPx = scaledFontSizePx(cellStyleLogicalFontSizeBasePx(st), viewZoom);
    const lineH = Math.max(12, fontPx * 1.25);
    const colW = sheet.getColWidth(c);
    const indentLv = clampIndentLevel(st?.indentLevel);
    const indentUnit = Math.max(6, fontPx * 0.55);
    let indentPx = indentLv * indentUnit;
    const maxIndentPx = Math.max(0, colW - 2 * PAD - 4);
    if (indentPx > maxIndentPx) {
      indentPx = maxIndentPx;
    }
    const innerW = Math.max(1, colW - 2 * PAD - indentPx);
    const lines = st?.wrapText === true ? wrapCellLines(ctx, text, innerW) : text.split("\n");
    const cellH = Math.ceil(lines.length * lineH + 4);
    maxH = Math.max(maxH, cellH);
  }
  return Math.max(MIN_LINE_SIZE, maxH > 0 ? maxH : sheet.defaultRowHeight);
}
