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
    const rowH = scaledRowHeightAt(sheet, r, env.viewZoom);
    for (let c = c0; c <= c1; c++) {
      const colW = scaledColWidthAt(sheet, c, env.viewZoom);
      if (colW <= 0 || rowH <= 0) continue;
      const x = cellLeftX(sheet, layout, c, env.viewZoom, env.scrollX);
      const y = cellTopY(sheet, layout, r, env.viewZoom, env.scrollY);
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
    const rowH = scaledRowHeightAt(sheet, r, env.viewZoom);
    for (let c = c0; c <= c1; c++) {
      const colW = scaledColWidthAt(sheet, c, env.viewZoom);
      if (colW <= 0 || rowH <= 0) continue;
      const x = cellLeftX(sheet, layout, c, env.viewZoom, env.scrollX);
      const y = cellTopY(sheet, layout, r, env.viewZoom, env.scrollY);
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

  const yTop = cellTopY(sheet, layout, r0, env.viewZoom, env.scrollY);
  const yBottom =
    cellTopY(sheet, layout, r1, env.viewZoom, env.scrollY) +
    scaledRowHeightAt(sheet, r1, env.viewZoom);
  const xLeft = cellLeftX(sheet, layout, c0, env.viewZoom, env.scrollX);
  const xRight =
    cellLeftX(sheet, layout, c1, env.viewZoom, env.scrollX) +
    scaledColWidthAt(sheet, c1, env.viewZoom);

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
    const y1 = snapLine(Math.max(by0, yTop));
    const y2 = snapLine(Math.min(by1, yBottom));
    if (y2 > y1) {
      ctx.moveTo(xs, y1);
      ctx.lineTo(xs, y2);
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
    const x1 = snapLine(Math.max(bx0, xLeft));
    const x2 = snapLine(Math.min(bx1, xRight));
    if (x2 > x1) {
      ctx.moveTo(x1, ys);
      ctx.lineTo(x2, ys);
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
