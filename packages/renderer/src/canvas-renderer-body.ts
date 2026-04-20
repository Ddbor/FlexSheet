import {
  formatCellDisplayWithStyle,
  type CellHorizontalAlign,
  type CellStyle,
  type CellTextOrientation,
  type CellVerticalAlign,
  type ConditionalFormattingOverlay,
  type Worksheet,
} from "@flexsheet/core";
import type { SheetTheme } from "@flexsheet/theme";
import { cellIntersectsCanvas, cellLeftX, cellTopY } from "./canvas-renderer-geometry.js";
import { paintCellFillPatternOverlay } from "./canvas-cell-fill-pattern.js";
import {
  argbToCss,
  buildCellCanvasFontWithLogicalPx,
  cellStyleLogicalFontSizeBasePx,
  paintAutoFilterDropdownGlyph,
  scaledColWidthAt,
  scaledFontSizePx,
  scaledRowHeightAt,
  snapLine,
  wrapCellLines,
} from "./canvas-renderer-utils.js";
import { paintBodyCellBorders } from "./canvas-renderer-body-borders.js";
import { getConditionalFormattingCellOverlayCached } from "./canvas-renderer-cf-overlay.js";
import { collectFrozenBodyQuadrantPasses, type BodyQuadrantPass } from "./frozen-body-quadrants.js";
import {
  COLUMN_HEADER_FILTER_BUTTON_CSS_PX,
  bodyColumnAutoFilterTextReservePx,
} from "./grid-hit-test.js";
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
  cfOverlayCellCache: Map<string, ConditionalFormattingOverlay | null> | undefined,
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
      const cfFill = getConditionalFormattingCellOverlayCached(
        sheet,
        r,
        c,
        cfOverlayCellCache,
      )?.fillArgb;
      const fillArgb = cfFill !== undefined && cfFill !== "" ? cfFill : cell.style?.fillArgb;
      const fillCss =
        fillArgb !== undefined && fillArgb !== ""
          ? (argbToCss(fillArgb) ?? env.theme.cellBg)
          : env.theme.cellBg;
      ctx.fillStyle = fillCss;
      ctx.fillRect(x, y, colW, rowH);
      const pat = cell.style?.fillPatternType ?? "none";
      if (
        (cfFill === undefined || cfFill === "") &&
        pat !== "none" &&
        cell.style !== undefined &&
        cell.style !== null
      ) {
        const rawFg = cell.style.fillPatternFgArgb?.trim();
        const fgArgb =
          rawFg !== undefined && rawFg !== "" && /^[\dA-Fa-f]{8}$/i.test(rawFg)
            ? rawFg.toUpperCase()
            : "FF000000";
        const fgCss = argbToCss(fgArgb) ?? "#000000";
        paintCellFillPatternOverlay(ctx, x, y, colW, rowH, pat, fgCss);
      }
    }
  }
}

function argbLerpTowardWhite(argbStr: string, t: number): string {
  const t0 = argbStr.trim();
  if (!/^[\dA-Fa-f]{8}$/i.test(t0)) {
    return "#4488cc";
  }
  const r = parseInt(t0.slice(2, 4), 16);
  const g = parseInt(t0.slice(4, 6), 16);
  const b = parseInt(t0.slice(6, 8), 16);
  const u = Math.max(0, Math.min(1, t));
  const rr = Math.round(r + (255 - r) * u);
  const gg = Math.round(g + (255 - g) * u);
  const bb = Math.round(b + (255 - b) * u);
  return `rgb(${rr},${gg},${bb})`;
}

function paintBodyCellDataBars(
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
  cfOverlayCellCache: Map<string, ConditionalFormattingOverlay | null> | undefined,
): void {
  const { ctx } = env;
  const padX = 3;
  const padY = 2;
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
      if (colW <= 0 || rowH <= 0) {
        continue;
      }
      if (!cellIntersectsCanvas(x, y, colW, rowH, headerW, headerH, canvasW, canvasH)) {
        continue;
      }
      const db = getConditionalFormattingCellOverlayCached(
        sheet,
        r,
        c,
        cfOverlayCellCache,
      )?.dataBar;
      if (db === undefined) {
        continue;
      }
      const innerW = colW - 2 * padX;
      const innerH = rowH - 2 * padY;
      if (innerW <= 1 || innerH <= 3) {
        continue;
      }
      const barH = Math.max(2, innerH * db.barHeightFrac);
      const top = y + padY + (innerH - barH) / 2;
      const left = x + padX;
      const x0 = left + db.barX0Frac * innerW;
      const x1 = left + db.barX1Frac * innerW;
      const w = Math.max(0, x1 - x0);
      if (w < 0.5) {
        continue;
      }
      const baseArgb = db.usePositiveFill ? db.posFillArgb : db.negFillArgb;
      const baseCss = argbToCss(baseArgb) ?? "#4488cc";

      ctx.save();
      ctx.beginPath();
      ctx.rect(left, top, innerW, barH);
      ctx.clip();

      if (db.fillKind === "gradient") {
        const g = ctx.createLinearGradient(x0, top, x1, top + barH);
        g.addColorStop(0, baseCss);
        g.addColorStop(1, argbLerpTowardWhite(baseArgb, 0.45));
        ctx.fillStyle = g;
      } else {
        ctx.fillStyle = baseCss;
      }
      ctx.fillRect(x0, top, w, barH);

      if (db.border.kind === "solid") {
        const bArgb = db.usePositiveFill ? db.border.posArgb : db.border.negArgb;
        const bCss = argbToCss(bArgb) ?? "#000000";
        ctx.strokeStyle = bCss;
        ctx.lineWidth = 1;
        ctx.strokeRect(x0 + 0.5, top + 0.5, w - 1, barH - 1);
      }

      if (db.axisXFrac !== null) {
        const ax = left + db.axisXFrac * innerW;
        const xs = snapLine(ax);
        ctx.strokeStyle = argbToCss(db.axisColorArgb) ?? "#000000";
        ctx.lineWidth = 1;
        ctx.beginPath();
        ctx.moveTo(xs, top);
        ctx.lineTo(xs, top + barH);
        ctx.stroke();
      }
      ctx.restore();
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

function strokeCellTextStrikethrough(
  ctx: CanvasRenderingContext2D,
  text: string,
  leftX: number,
  baselineY: number,
  kind: CanvasTextBaseline,
  color: string,
  fontPx: number,
): void {
  const m = ctx.measureText(text);
  const w = m.width;
  let yStrike: number;
  if (kind === "middle") {
    yStrike = baselineY;
  } else {
    const ascent = m.actualBoundingBoxAscent ?? fontPx * 0.72;
    yStrike = baselineY - ascent * 0.45;
  }
  ctx.save();
  ctx.strokeStyle = color;
  ctx.lineWidth = Math.max(1, fontPx / 14);
  ctx.beginPath();
  ctx.moveTo(leftX, yStrike);
  ctx.lineTo(leftX + w, yStrike);
  ctx.stroke();
  ctx.restore();
}

function resolvedHAlign(st: CellStyle | null | undefined): CellHorizontalAlign {
  const h = st?.hAlign as CellHorizontalAlign | undefined;
  if (
    h === "center" ||
    h === "right" ||
    h === "fill" ||
    h === "justify" ||
    h === "distributed" ||
    h === "centerContinuous"
  ) {
    return h;
  }
  return "left";
}

function resolvedVAlign(st: CellStyle | null | undefined): CellVerticalAlign {
  const v = st?.vAlign as CellVerticalAlign | undefined;
  if (v === "top" || v === "bottom" || v === "justify" || v === "distributed") {
    return v;
  }
  return "middle";
}

function resolvedRotationDegrees(st: CellStyle | null | undefined): number | null {
  const d = st?.textRotationDegrees;
  if (d === undefined || !Number.isFinite(d) || d === 0) {
    return null;
  }
  return Math.max(-90, Math.min(90, d));
}

function resolvedTextOrientation(st: CellStyle | null | undefined): CellTextOrientation {
  return st?.textOrientation ?? "horizontal";
}

/** 非水平方向时忽略下划线（与 Canvas 测量一致）。 */
function paintOrientedBodyText(
  ctx: CanvasRenderingContext2D,
  text: string,
  x: number,
  y: number,
  drawW: number,
  rowH: number,
  orient: Exclude<CellTextOrientation, "horizontal">,
  fontPx: number,
): void {
  const cx = x + drawW / 2;
  const cy = y + rowH / 2;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  if (orient === "verticalStack") {
    const chars = Array.from(text);
    const lineH = Math.max(fontPx * 1.2, 12);
    const totalH = chars.length * lineH;
    let yy = cy - totalH / 2 + lineH / 2;
    for (const ch of chars) {
      if (yy - lineH / 2 > y + rowH + 1e-6) {
        break;
      }
      ctx.fillText(ch, cx, yy);
      yy += lineH;
    }
    return;
  }
  ctx.save();
  ctx.translate(cx, cy);
  switch (orient) {
    case "angleUp45":
      ctx.rotate(-Math.PI / 4);
      break;
    case "angleDown45":
      ctx.rotate(Math.PI / 4);
      break;
    case "rotateUp90":
      ctx.rotate(-Math.PI / 2);
      break;
    case "rotateDown90":
      ctx.rotate(Math.PI / 2);
      break;
    default:
      break;
  }
  ctx.fillText(text, 0, 0);
  ctx.restore();
}

function clampIndentLevel(raw: number | undefined): number {
  if (raw === undefined || !Number.isFinite(raw)) {
    return 0;
  }
  return Math.max(0, Math.min(255, Math.round(raw)));
}

function textLineLeftInBox(
  hAlign: CellHorizontalAlign,
  boxLeft: number,
  innerW: number,
  lineWidth: number,
): number {
  if (hAlign === "right") {
    return boxLeft + innerW - lineWidth;
  }
  if (hAlign === "center" || hAlign === "centerContinuous") {
    return boxLeft + (innerW - lineWidth) / 2;
  }
  return boxLeft;
}

function paintArbitraryRotatedText(
  ctx: CanvasRenderingContext2D,
  text: string,
  x: number,
  y: number,
  drawW: number,
  rowH: number,
  degrees: number,
  _fontPx: number,
): void {
  const cx = x + drawW / 2;
  const cy = y + rowH / 2;
  ctx.textAlign = "center";
  ctx.textBaseline = "middle";
  ctx.save();
  ctx.translate(cx, cy);
  /** OOXML / Excel：textRotation 逆时针为正；Canvas 2D 正角为顺时针，故取反与单元格一致。 */
  ctx.rotate((-degrees * Math.PI) / 180);
  ctx.fillText(text, 0, 0);
  ctx.restore();
}

function buildFillCellText(ctx: CanvasRenderingContext2D, text: string, innerW: number): string {
  if (text === "" || innerW <= 1) {
    return text;
  }
  const w0 = ctx.measureText(text).width;
  if (w0 <= 0) {
    return text;
  }
  let out = text;
  while (ctx.measureText(out + text).width <= innerW) {
    out += text;
  }
  while (out.length > text.length && ctx.measureText(out).width > innerW) {
    out = out.slice(0, -1);
  }
  return out;
}

function paintDistributedLine(
  ctx: CanvasRenderingContext2D,
  text: string,
  boxLeft: number,
  innerW: number,
  baselineY: number,
): void {
  const chars = Array.from(text);
  if (chars.length === 0) {
    return;
  }
  if (chars.length === 1) {
    const ch = chars[0]!;
    const w = ctx.measureText(ch).width;
    ctx.fillText(ch, boxLeft + (innerW - w) / 2, baselineY);
    return;
  }
  let total = 0;
  const widths = chars.map((ch) => {
    const w = ctx.measureText(ch).width;
    total += w;
    return w;
  });
  const gap = (innerW - total) / (chars.length - 1);
  if (gap < 0) {
    ctx.fillText(text, boxLeft, baselineY);
    return;
  }
  let xx = boxLeft;
  for (let i = 0; i < chars.length; i++) {
    const ch = chars[i]!;
    ctx.fillText(ch, xx, baselineY);
    xx += widths[i]! + gap;
  }
}

function paintJustifiedLine(
  ctx: CanvasRenderingContext2D,
  line: string,
  boxLeft: number,
  innerW: number,
  baselineY: number,
): void {
  const parts = line.trim().split(/\s+/).filter((s) => s.length > 0);
  if (parts.length <= 1) {
    ctx.fillText(line, boxLeft, baselineY);
    return;
  }
  let total = 0;
  const ws = parts.map((w) => {
    const m = ctx.measureText(w).width;
    total += m;
    return m;
  });
  const spaceW = ctx.measureText(" ").width;
  const gaps = parts.length - 1;
  const extra = innerW - total - gaps * spaceW;
  const add = gaps > 0 ? extra / gaps : 0;
  let xx = boxLeft;
  for (let i = 0; i < parts.length; i++) {
    const w = parts[i]!;
    ctx.fillText(w, xx, baselineY);
    xx += ws[i]!;
    if (i < parts.length - 1) {
      xx += spaceW + add;
    }
  }
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
  cfOverlayCellCache: Map<string, ConditionalFormattingOverlay | null> | undefined,
): void {
  const { ctx } = env;
  const hasVisibleContent = (row: number, col: number): boolean => {
    if (col < 0 || col >= sheet.colCount || row < 0 || row >= sheet.rowCount) {
      return false;
    }
    const cell = sheet.getCell(row, col);
    return formatCellDisplayWithStyle(cell.value, cell.style) !== "";
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
      const text = formatCellDisplayWithStyle(cell.value, cell.style);
      if (text === "") {
        continue;
      }
      const cfOv = getConditionalFormattingCellOverlayCached(sheet, r, c, cfOverlayCellCache);
      if (cfOv?.dataBar?.hideCellValue === true) {
        continue;
      }
      const cfFg = cfOv?.fgArgb;
      const fgArgb = cfFg !== undefined && cfFg !== "" ? cfFg : cell.style?.fgArgb;
      ctx.fillStyle =
        fgArgb !== undefined && fgArgb !== ""
          ? (argbToCss(fgArgb) ?? env.theme.cellColor)
          : env.theme.cellColor;
      ctx.textAlign = "left";
      const pad = 4;
      const filterReservePx =
        bodyColumnAutoFilterTextReservePx(sheet, r, c) + sheet.getPivotPageFilterDropdownReservePx(r, c);
      const hAlign = resolvedHAlign(cell.style);
      const vAlign = resolvedVAlign(cell.style);
      const baseLog = cellStyleLogicalFontSizeBasePx(cell.style);
      const baseFontPx = scaledFontSizePx(baseLog, env.viewZoom);
      const indentLv = clampIndentLevel(cell.style?.indentLevel);
      const indentUnit = Math.max(6, baseFontPx * 0.55);
      let indentPx = indentLv * indentUnit;
      const maxIndentPx = Math.max(0, colW - 2 * pad - 4 - filterReservePx);
      if (indentPx > maxIndentPx) {
        indentPx = maxIndentPx;
      }
      const boxLeft = x + pad + indentPx;
      const innerW = Math.max(1, colW - 2 * pad - indentPx - filterReservePx);

      const wrap = cell.style?.wrapText === true;
      const orient = resolvedTextOrientation(cell.style);
      const rotDeg = resolvedRotationDegrees(cell.style);
      const isVerticalStack = orient === "verticalStack";
      const fixedOrientActive = !isVerticalStack && orient !== "horizontal" && rotDeg === null;
      const rotAngleActive = rotDeg !== null;

      ctx.font = buildCellCanvasFontWithLogicalPx(cell.style, baseLog, env.viewZoom);
      let lines: string[];
      if (isVerticalStack || fixedOrientActive) {
        const raw = text.includes("\n") ? (text.split("\n")[0] ?? "") : text;
        lines = [raw];
      } else if (rotAngleActive) {
        const raw = text.includes("\n") ? (text.split("\n")[0] ?? "") : text;
        lines = [raw];
      } else if (wrap) {
        lines = wrapCellLines(ctx, text, innerW);
      } else if (text.includes("\n")) {
        lines = text.split("\n");
      } else {
        lines = [text];
      }

      let logicalFontPx = baseLog;
      if (
        cell.style?.shrinkToFit === true &&
        !wrap &&
        !isVerticalStack &&
        !fixedOrientActive &&
        !rotAngleActive
      ) {
        let log = baseLog;
        const minLog = Math.max(4, baseLog * 0.35);
        const probe = lines[0] ?? text;
        while (log >= minLog) {
          ctx.font = buildCellCanvasFontWithLogicalPx(cell.style, log, env.viewZoom);
          if (probe === "" || ctx.measureText(probe).width <= innerW) {
            break;
          }
          log -= 0.5;
        }
        logicalFontPx = log;
      }
      ctx.font = buildCellCanvasFontWithLogicalPx(cell.style, logicalFontPx, env.viewZoom);
      const fontPx = scaledFontSizePx(logicalFontPx, env.viewZoom);

      const onlySingleLineVisual =
        isVerticalStack || fixedOrientActive || rotAngleActive || lines.length === 1;
      const drawW =
        onlySingleLineVisual && !fixedOrientActive && !isVerticalStack && !rotAngleActive
          ? extendedTextWidth(r, c, colW)
          : colW;
      const underlineKind = cell.style?.underline;
      const ink = String(ctx.fillStyle);
      ctx.save();
      ctx.beginPath();
      if (filterReservePx > 0) {
        const wAnchor = Math.max(0, colW - filterReservePx);
        ctx.rect(x, y, wAnchor, rowH);
        if (drawW > colW) {
          ctx.rect(x + colW, y, drawW - colW, rowH);
        }
      } else {
        ctx.rect(x, y, drawW, rowH);
      }
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
          if (hAlign === "justify" && li < lines.length - 1) {
            paintJustifiedLine(ctx, line, boxLeft, innerW, baselineY);
          } else if (hAlign === "distributed") {
            paintDistributedLine(ctx, line, boxLeft, innerW, baselineY);
          } else {
            ctx.fillText(line, lineLeft, baselineY);
          }
          if (
            (underlineKind === "single" || underlineKind === "double") &&
            hAlign !== "justify" &&
            hAlign !== "distributed"
          ) {
            strokeCellTextUnderline(ctx, line, lineLeft, baselineY, underlineKind, ink, fontPx);
          }
          if (
            cell.style?.strikethrough === true &&
            hAlign !== "justify" &&
            hAlign !== "distributed"
          ) {
            strokeCellTextStrikethrough(ctx, line, lineLeft, baselineY, "alphabetic", ink, fontPx);
          }
          yy += lineH;
        }
      } else if (isVerticalStack || fixedOrientActive) {
        const line = lines[0] ?? "";
        if (line !== "") {
          paintOrientedBodyText(
            ctx,
            line,
            x,
            y,
            drawW,
            rowH,
            orient as Exclude<CellTextOrientation, "horizontal">,
            fontPx,
          );
        }
      } else if (rotAngleActive && rotDeg !== null) {
        const line = lines[0] ?? "";
        if (line !== "") {
          paintArbitraryRotatedText(ctx, line, x, y, drawW, rowH, rotDeg, fontPx);
        }
      } else {
        const line = lines[0] ?? "";
        let display = line;
        if (hAlign === "fill" && !wrap) {
          display = buildFillCellText(ctx, line, innerW);
        }
        const scriptKind = cell.style?.fontScript;
        const useScript =
          (scriptKind === "superscript" || scriptKind === "subscript") && hAlign !== "distributed";
        const logicalDraw = useScript ? logicalFontPx * 0.65 : logicalFontPx;
        ctx.font = buildCellCanvasFontWithLogicalPx(cell.style, logicalDraw, env.viewZoom);
        const fontPxDraw = scaledFontSizePx(logicalDraw, env.viewZoom);
        const m = ctx.measureText(display);
        const ascent = m.actualBoundingBoxAscent ?? fontPxDraw * 0.72;
        const descent = m.actualBoundingBoxDescent ?? fontPxDraw * 0.22;
        const textLeft = textLineLeftInBox(hAlign, boxLeft, innerW, m.width);
        let baselineY: number;
        let tb: CanvasTextBaseline;
        if (vAlign === "top") {
          tb = "alphabetic";
          ctx.textBaseline = "alphabetic";
          baselineY = y + pad + ascent;
        } else if (vAlign === "bottom") {
          tb = "alphabetic";
          ctx.textBaseline = "alphabetic";
          baselineY = y + rowH - pad - descent;
        } else {
          tb = "middle";
          ctx.textBaseline = "middle";
          baselineY = y + rowH / 2;
        }
        if (useScript && scriptKind !== undefined) {
          const basePx = scaledFontSizePx(logicalFontPx, env.viewZoom);
          baselineY += scriptKind === "superscript" ? -basePx * 0.28 : basePx * 0.14;
        }
        if (hAlign === "distributed") {
          paintDistributedLine(ctx, display, boxLeft, innerW, baselineY);
        } else {
          ctx.fillText(display, textLeft, baselineY);
        }
        if (
          (underlineKind === "single" || underlineKind === "double") &&
          hAlign !== "distributed"
        ) {
          strokeCellTextUnderline(ctx, display, textLeft, baselineY, underlineKind, ink, fontPxDraw);
        }
        if (
          cell.style?.strikethrough === true &&
          hAlign !== "distributed" &&
          hAlign !== "justify"
        ) {
          strokeCellTextStrikethrough(ctx, display, textLeft, baselineY, tb, ink, fontPxDraw);
        }
      }
      ctx.restore();
    }
  }
}

/** 透视表页字段「值」格右侧下拉按钮（与列筛选同形）。 */
function paintPivotPageFilterAnchors(
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
  for (const p of sheet.getPivotTableDefinitionsSnapshot()) {
    const fc = p.filterFieldCols.length;
    if (fc === 0) {
      continue;
    }
    const base = p.pageFilterStartRow ?? p.destinationRow;
    const dc = p.destinationCol;
    for (let i = 0; i < fc; i++) {
      const r = base + i;
      const c = dc + 1;
      if (r < r0 || r > r1 || c < c0 || c > c1) {
        continue;
      }
      if (sheet.isMergeCoveredCell(r, c)) {
        continue;
      }
      const anchor = sheet.getMergeAnchorCell(r, c);
      if (anchor.row !== r || anchor.col !== c) {
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
      if (colW <= 0 || rowH <= 0) {
        continue;
      }
      if (!cellIntersectsCanvas(x, y, colW, rowH, headerW, headerH, canvasW, canvasH)) {
        continue;
      }
      const pad = 2;
      const innerW = COLUMN_HEADER_FILTER_BUTTON_CSS_PX - 4;
      const innerH = 14;
      const bx = x + colW - COLUMN_HEADER_FILTER_BUTTON_CSS_PX - pad;
      const by = y + Math.max(1, (Math.min(rowH, 22) - innerH) / 2);
      paintAutoFilterDropdownGlyph(ctx, bx, by, innerW, innerH, {
        narrowed: false,
        sortHint: null,
        borderColor: env.theme.gridLineColor,
      });
    }
  }
}

function paintBodyAutoFilterAnchors(
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
      const meta = sheet.getColumnAutoFilterMeta(c);
      if (meta?.uiKind !== "body" || meta.bodyAnchorRow !== r) {
        continue;
      }
      const anchor = sheet.getMergeAnchorCell(r, c);
      if (anchor.row !== r || anchor.col !== c) {
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
      if (colW <= 0 || rowH <= 0) {
        continue;
      }
      if (!cellIntersectsCanvas(x, y, colW, rowH, headerW, headerH, canvasW, canvasH)) {
        continue;
      }
      const pad = 2;
      const innerW = COLUMN_HEADER_FILTER_BUTTON_CSS_PX - 4;
      const innerH = 14;
      const bx = x + colW - COLUMN_HEADER_FILTER_BUTTON_CSS_PX - pad;
      const by = y + Math.max(1, (Math.min(rowH, 22) - innerH) / 2);
      paintAutoFilterDropdownGlyph(ctx, bx, by, innerW, innerH, {
        narrowed: sheet.isColumnAutoFilterNarrowed(c),
        sortHint: sheet.getColumnAutoFilterSortHint(c) ?? null,
        borderColor: env.theme.gridLineColor,
      });
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
  const cfOverlayCellCache =
    sheet.getConditionalFormatRules().length > 0
      ? new Map<string, ConditionalFormattingOverlay | null>()
      : undefined;
  paintBodyCellFills(
    env,
    sheet,
    layout,
    headerW,
    headerH,
    canvasW,
    canvasH,
    r0,
    r1,
    c0,
    c1,
    cfOverlayCellCache,
  );
  paintBodyCellDataBars(
    env,
    sheet,
    layout,
    headerW,
    headerH,
    canvasW,
    canvasH,
    r0,
    r1,
    c0,
    c1,
    cfOverlayCellCache,
  );

  if (strokeBounds !== null && env.showGridLines) {
    strokeBodyGrid(env, sheet, layout, r0, r1, c0, c1, strokeBounds);
  }

  paintBodyCellBorders(
    env,
    sheet,
    layout,
    headerW,
    headerH,
    canvasW,
    canvasH,
    r0,
    r1,
    c0,
    c1,
    cfOverlayCellCache,
  );
  paintBodyCellTexts(
    env,
    sheet,
    layout,
    headerW,
    headerH,
    canvasW,
    canvasH,
    r0,
    r1,
    c0,
    c1,
    cfOverlayCellCache,
  );
  paintBodyAutoFilterAnchors(
    env,
    sheet,
    layout,
    headerW,
    headerH,
    canvasW,
    canvasH,
    r0,
    r1,
    c0,
    c1,
  );
  paintPivotPageFilterAnchors(
    env,
    sheet,
    layout,
    headerW,
    headerH,
    canvasW,
    canvasH,
    r0,
    r1,
    c0,
    c1,
  );
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
