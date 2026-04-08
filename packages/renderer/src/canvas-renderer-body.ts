import type { CellStyle, Worksheet } from "@flexsheet/core";
import type { SheetTheme } from "@flexsheet/theme";
import { cellIntersectsCanvas, cellLeftX, cellTopY } from "./canvas-renderer-geometry.js";
import {
  argbToCss,
  buildCellCanvasFont,
  cellStyleLogicalFontSizeBasePx,
  formatCellDisplay,
  scaledColWidthAt,
  scaledFontSizePx,
  scaledRowHeightAt,
  snapLine,
  wrapCellLines,
} from "./canvas-renderer-utils.js";
import { paintBodyCellBorders } from "./canvas-renderer-body-borders.js";
import { collectFrozenBodyQuadrantPasses, type BodyQuadrantPass } from "./frozen-body-quadrants.js";
import type { FrozenLayout } from "./viewport.js";

export interface BodyPaintEnv {
  readonly ctx: CanvasRenderingContext2D;
  readonly theme: SheetTheme;
  readonly viewZoom: number;
  readonly showGridLines: boolean;
  readonly scrollX: number;
  readonly scrollY: number;
  readonly viewportBuffer: number;
}

function paintBodyCellFills(
  env: BodyPaintEnv,
  sheet: Worksheet,
  layout: FrozenLayout,
  headerW: number,
  headerH: number,
  canvasW: number,
  canvasH: number,
  r0: number,
  r1: number,
  c0: number,
  c1: number,
): void {
  const { ctx } = env;
  for (let r = r0; r <= r1; r++) {
    for (let c = c0; c <= c1; c++) {
      if (sheet.isMergeCoveredCell(r, c)) {
        continue;
      }
      const info = sheet.getMergedRectInfo(r, c);
      let colW = scaledColWidthAt(sheet, c, env.viewZoom);
      let rowH = scaledRowHeightAt(sheet, r, env.viewZoom);
      let x = cellLeftX(sheet, layout, c, env.viewZoom, env.scrollX);
      let y = cellTopY(sheet, layout, r, env.viewZoom, env.scrollY);
      if (info.rowSpan > 1 || info.colSpan > 1) {
        colW = 0;
        for (let cc = info.anchorCol; cc < info.anchorCol + info.colSpan; cc++) {
          colW += scaledColWidthAt(sheet, cc, env.viewZoom);
        }
        rowH = 0;
        for (let rr = info.anchorRow; rr < info.anchorRow + info.rowSpan; rr++) {
          rowH += scaledRowHeightAt(sheet, rr, env.viewZoom);
        }
        x = cellLeftX(sheet, layout, info.anchorCol, env.viewZoom, env.scrollX);
        y = cellTopY(sheet, layout, info.anchorRow, env.viewZoom, env.scrollY);
      }
      if (colW <= 0 || rowH <= 0) continue;
      if (!cellIntersectsCanvas(x, y, colW, rowH, headerW, headerH, canvasW, canvasH)) {
        continue;
      }
      const cell = sheet.getCell(r, c);
      const fillArgb = cell.style?.fillArgb;
      const fillCss =
        fillArgb !== undefined && fillArgb !== ""
          ? (argbToCss(fillArgb) ?? env.theme.cellBg)
          : env.theme.cellBg;
      ctx.fillStyle = fillCss;
      ctx.fillRect(x, y, colW, rowH);
    }
  }
}

function strokeCellTextUnderline(
  ctx: CanvasRenderingContext2D,
  text: string,
  leftX: number,
  baselineY: number,
  kind: "single" | "double",
  color: string,
  fontPx: number,
): void {
  const m = ctx.measureText(text);
  const descent = m.actualBoundingBoxDescent ?? Math.max(2, fontPx * 0.22);
  const y1 = baselineY + descent + 1;
  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = Math.max(1, fontPx / 14);
  ctx.beginPath();
  ctx.moveTo(leftX, y1);
  ctx.lineTo(leftX + m.width, y1);
  ctx.stroke();
  if (kind === "double") {
    const gap = Math.max(2, fontPx * 0.12);
    const y2 = y1 + gap;
    ctx.beginPath();
    ctx.moveTo(leftX, y2);
    ctx.lineTo(leftX + m.width, y2);
    ctx.stroke();
  }
  ctx.restore();
}

function resolvedHAlign(st: CellStyle | null | undefined): "left" | "center" | "right" {
  const h = st?.hAlign;
  return h === "center" || h === "right" ? h : "left";
}

function resolvedVAlign(st: CellStyle | null | undefined): "top" | "middle" | "bottom" {
  const v = st?.vAlign;
  return v === "top" || v === "bottom" ? v : "middle";
}

function clampIndentLevel(raw: number | undefined): number {
  if (raw === undefined || !Number.isFinite(raw)) {
    return 0;
  }
  return Math.max(0, Math.min(255, Math.round(raw)));
}

function textLineLeftInBox(
  hAlign: "left" | "center" | "right",
  boxLeft: number,
  innerW: number,
  lineWidth: number,
): number {
  if (hAlign === "center") {
    return boxLeft + (innerW - lineWidth) / 2;
  }
  if (hAlign === "right") {
    return boxLeft + innerW - lineWidth;
  }
  return boxLeft;
}

function paintBodyCellTexts(
  env: BodyPaintEnv,
  sheet: Worksheet,
  layout: FrozenLayout,
  headerW: number,
  headerH: number,
  canvasW: number,
  canvasH: number,
  r0: number,
  r1: number,
  c0: number,
  c1: number,
): void {
  const { ctx } = env;
  const hasVisibleContent = (row: number, col: number): boolean => {
    if (col < 0 || col >= sheet.colCount || row < 0 || row >= sheet.rowCount) {
      return false;
    }
    return formatCellDisplay(sheet.getCell(row, col).value) !== "";
  };

  const extendedTextWidth = (row: number, col: number, baseW: number): number => {
    const m = sheet.getMergedRectInfo(row, col);
    if (m.colSpan > 1) {
      return baseW;
    }
    let w = baseW;
    for (let c = col + 1; c < sheet.colCount; c++) {
      if (hasVisibleContent(row, c)) {
        break;
      }
      w += scaledColWidthAt(sheet, c, env.viewZoom);
    }
    return w;
  };

  for (let r = r0; r <= r1; r++) {
    for (let c = c0; c <= c1; c++) {
      if (sheet.isMergeCoveredCell(r, c)) {
        continue;
      }
      const info = sheet.getMergedRectInfo(r, c);
      let colW = scaledColWidthAt(sheet, c, env.viewZoom);
      let rowH = scaledRowHeightAt(sheet, r, env.viewZoom);
      let x = cellLeftX(sheet, layout, c, env.viewZoom, env.scrollX);
      let y = cellTopY(sheet, layout, r, env.viewZoom, env.scrollY);
      if (info.rowSpan > 1 || info.colSpan > 1) {
        colW = 0;
        for (let cc = info.anchorCol; cc < info.anchorCol + info.colSpan; cc++) {
          colW += scaledColWidthAt(sheet, cc, env.viewZoom);
        }
        rowH = 0;
        for (let rr = info.anchorRow; rr < info.anchorRow + info.rowSpan; rr++) {
          rowH += scaledRowHeightAt(sheet, rr, env.viewZoom);
        }
        x = cellLeftX(sheet, layout, info.anchorCol, env.viewZoom, env.scrollX);
        y = cellTopY(sheet, layout, info.anchorRow, env.viewZoom, env.scrollY);
      }
      if (colW <= 0 || rowH <= 0) continue;
      if (!cellIntersectsCanvas(x, y, colW, rowH, headerW, headerH, canvasW, canvasH)) {
        continue;
      }
      const cell = sheet.getCell(r, c);
      const text = formatCellDisplay(cell.value);
      if (text === "") {
        continue;
      }
      const fgArgb = cell.style?.fgArgb;
      ctx.fillStyle =
        fgArgb !== undefined && fgArgb !== ""
          ? (argbToCss(fgArgb) ?? env.theme.cellColor)
          : env.theme.cellColor;
      ctx.font = buildCellCanvasFont(cell.style, env.viewZoom);
      const fontPx = scaledFontSizePx(cellStyleLogicalFontSizeBasePx(cell.style), env.viewZoom);
      ctx.textAlign = "left";
      const pad = 4;
      const hAlign = resolvedHAlign(cell.style);
      const vAlign = resolvedVAlign(cell.style);
      const indentLv = clampIndentLevel(cell.style?.indentLevel);
      const indentUnit = Math.max(6, fontPx * 0.55);
      let indentPx = indentLv * indentUnit;
      const maxIndentPx = Math.max(0, colW - 2 * pad - 4);
      if (indentPx > maxIndentPx) {
        indentPx = maxIndentPx;
      }
      const boxLeft = x + pad + indentPx;
      const innerW = Math.max(1, colW - 2 * pad - indentPx);

      const wrap = cell.style?.wrapText === true;
      let lines: string[];
      if (wrap) {
        lines = wrapCellLines(ctx, text, innerW);
      } else if (text.includes("\n")) {
        lines = text.split("\n");
      } else {
        lines = [text];
      }

      const onlySingleLineVisual = !wrap && !text.includes("\n");
      const drawW = onlySingleLineVisual ? extendedTextWidth(r, c, colW) : colW;
      const underlineKind = cell.style?.underline;
      const ink = String(ctx.fillStyle);
      ctx.save();
      ctx.beginPath();
      ctx.rect(x, y, drawW, rowH);
      ctx.clip();
      if (!onlySingleLineVisual) {
        const lineH = Math.max(12, fontPx * 1.25);
        const padY = 2;
        let fitCount = 0;
        let acc = padY;
        for (const _ of lines) {
          if (acc + lineH > y + rowH - padY + 1e-6) {
            break;
          }
          fitCount += 1;
          acc += lineH;
        }
        const n = Math.min(lines.length, fitCount);
        const blockH = n * lineH;
        let yy = y + padY;
        if (n > 0) {
          yy =
            vAlign === "top"
              ? y + padY
              : vAlign === "bottom"
                ? y + rowH - padY - blockH
                : y + (rowH - blockH) / 2;
          const yyMax = y + rowH - padY - blockH;
          const yyMin = y + padY;
          if (yy < yyMin) {
            yy = yyMin;
          }
          if (yy > yyMax) {
            yy = yyMax;
          }
        }
        for (let li = 0; li < lines.length; li++) {
          const line = lines[li]!;
          if (yy + lineH > y + rowH + 1e-6) {
            break;
          }
          const mLine = ctx.measureText(line);
          const ascent = mLine.actualBoundingBoxAscent ?? fontPx * 0.72;
          const lineW = mLine.width;
          const lineLeft = textLineLeftInBox(hAlign, boxLeft, innerW, lineW);
          ctx.textBaseline = "alphabetic";
          const baselineY = yy + ascent;
          ctx.fillText(line, lineLeft, baselineY);
          if (underlineKind === "single" || underlineKind === "double") {
            strokeCellTextUnderline(ctx, line, lineLeft, baselineY, underlineKind, ink, fontPx);
          }
          yy += lineH;
        }
      } else {
        const m = ctx.measureText(text);
        const ascent = m.actualBoundingBoxAscent ?? fontPx * 0.72;
        const descent = m.actualBoundingBoxDescent ?? fontPx * 0.22;
        const textLeft = textLineLeftInBox(hAlign, boxLeft, innerW, m.width);
        let baselineY: number;
        if (vAlign === "top") {
          ctx.textBaseline = "alphabetic";
          baselineY = y + pad + ascent;
        } else if (vAlign === "bottom") {
          ctx.textBaseline = "alphabetic";
          baselineY = y + rowH - pad - descent;
        } else {
          ctx.textBaseline = "middle";
          baselineY = y + rowH / 2;
        }
        ctx.fillText(text, textLeft, baselineY);
        if (underlineKind === "single" || underlineKind === "double") {
          strokeCellTextUnderline(ctx, text, textLeft, baselineY, underlineKind, ink, fontPx);
        }
      }
      ctx.restore();
    }
  }
}

function verticalMergeEdgeHidden(sheet: Worksheet, row: number, leftCol: number): boolean {
  if (leftCol < 0) {
    return false;
  }
  const a = sheet.getMergeAnchorCell(row, leftCol);
  const b = sheet.getMergeAnchorCell(row, leftCol + 1);
  return a.row === b.row && a.col === b.col;
}

function horizontalMergeEdgeHidden(sheet: Worksheet, topRow: number, col: number): boolean {
  if (topRow < 0) {
    return false;
  }
  const a = sheet.getMergeAnchorCell(topRow, col);
  const b = sheet.getMergeAnchorCell(topRow + 1, col);
  return a.row === b.row && a.col === b.col;
}

function strokeBodyGrid(
  env: BodyPaintEnv,
  sheet: Worksheet,
  layout: FrozenLayout,
  r0: number,
  r1: number,
  c0: number,
  c1: number,
  clipBounds: { minX: number; minY: number; maxX: number; maxY: number },
): void {
  const { ctx } = env;
  if (r0 > r1 || c0 > c1) {
    return;
  }

  const { minX: bx0, minY: by0, maxX: bx1, maxY: by1 } = clipBounds;
  const { frozenCols, frozenRows } = layout;

  ctx.strokeStyle = env.theme.gridLineColor;
  ctx.lineWidth = 1;
  ctx.beginPath();

  for (let c = c0; c <= c1 + 1; c++) {
    const x =
      c < sheet.colCount
        ? cellLeftX(sheet, layout, c, env.viewZoom, env.scrollX)
        : cellLeftX(sheet, layout, sheet.colCount - 1, env.viewZoom, env.scrollX) +
          scaledColWidthAt(sheet, sheet.colCount - 1, env.viewZoom);

    if (c === frozenCols) {
      continue;
    }

    if (x < bx0 || x > bx1) {
      continue;
    }
    const xs = snapLine(x);
    for (let r = r0; r <= r1; r++) {
      if (c > 0 && verticalMergeEdgeHidden(sheet, r, c - 1)) {
        continue;
      }
      const y1 = cellTopY(sheet, layout, r, env.viewZoom, env.scrollY);
      const y2 = y1 + scaledRowHeightAt(sheet, r, env.viewZoom);
      const segY1 = snapLine(Math.max(by0, y1));
      const segY2 = snapLine(Math.min(by1, y2));
      if (segY2 > segY1) {
        ctx.moveTo(xs, segY1);
        ctx.lineTo(xs, segY2);
      }
    }
  }

  for (let r = r0; r <= r1 + 1; r++) {
    const y =
      r < sheet.rowCount
        ? cellTopY(sheet, layout, r, env.viewZoom, env.scrollY)
        : cellTopY(sheet, layout, sheet.rowCount - 1, env.viewZoom, env.scrollY) +
          scaledRowHeightAt(sheet, sheet.rowCount - 1, env.viewZoom);

    if (r === frozenRows) {
      continue;
    }

    if (y < by0 || y > by1) {
      continue;
    }
    const ys = snapLine(y);
    for (let c = c0; c <= c1; c++) {
      if (r > 0 && horizontalMergeEdgeHidden(sheet, r - 1, c)) {
        continue;
      }
      const x1 = cellLeftX(sheet, layout, c, env.viewZoom, env.scrollX);
      const x2 = x1 + scaledColWidthAt(sheet, c, env.viewZoom);
      const segX1 = snapLine(Math.max(bx0, x1));
      const segX2 = snapLine(Math.min(bx1, x2));
      if (segX2 > segX1) {
        ctx.moveTo(segX1, ys);
        ctx.lineTo(segX2, ys);
      }
    }
  }

  ctx.stroke();
}

function runBodyQuadrantPass(
  env: BodyPaintEnv,
  sheet: Worksheet,
  layout: FrozenLayout,
  headerW: number,
  headerH: number,
  canvasW: number,
  canvasH: number,
  pass: BodyQuadrantPass,
): void {
  const { ctx } = env;
  const { clipX, clipY, clipW, clipH, r0, r1, c0, c1 } = pass;
  if (r0 > r1 || c0 > c1) {
    return;
  }
  let strokeBounds: { minX: number; minY: number; maxX: number; maxY: number } | null = null;
  if (clipW > 0 && clipH > 0) {
    strokeBounds = {
      minX: clipX,
      minY: clipY,
      maxX: clipX + clipW,
      maxY: clipY + clipH,
    };
  }
  ctx.save();
  ctx.beginPath();
  ctx.rect(clipX, clipY, clipW, clipH);
  ctx.clip();
  paintBodyCellFills(env, sheet, layout, headerW, headerH, canvasW, canvasH, r0, r1, c0, c1);

  if (strokeBounds !== null && env.showGridLines) {
    strokeBodyGrid(env, sheet, layout, r0, r1, c0, c1, strokeBounds);
  }

  paintBodyCellBorders(env, sheet, layout, headerW, headerH, canvasW, canvasH, r0, r1, c0, c1);
  paintBodyCellTexts(env, sheet, layout, headerW, headerH, canvasW, canvasH, r0, r1, c0, c1);
  ctx.restore();
}

export function drawBody(
  env: BodyPaintEnv,
  sheet: Worksheet,
  layout: FrozenLayout,
  headerW: number,
  headerH: number,
  canvasW: number,
  canvasH: number,
): void {
  const passes = collectFrozenBodyQuadrantPasses(
    sheet,
    layout,
    headerW,
    headerH,
    canvasW,
    canvasH,
    env.scrollX,
    env.scrollY,
    env.viewZoom,
    env.viewportBuffer,
  );
  for (const pass of passes) {
    runBodyQuadrantPass(env, sheet, layout, headerW, headerH, canvasW, canvasH, pass);
  }
}
