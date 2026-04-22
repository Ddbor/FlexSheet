import { isUnconfiguredPivotDefinition, pivotLayoutStartRow, type Worksheet } from "@flexsheet/core";
import { cellIntersectsCanvas, cellLeftX, cellTopY } from "./canvas-renderer-geometry.js";
import { scaledColWidthAt, scaledRowHeightAt } from "./canvas-renderer-utils.js";
import type { FrozenLayout } from "./viewport.js";

/**
 * 在未配置透视表输出区第 3 行起的区域绘制与 Excel 类似的示意图案（Canvas 矢量，非单元格图案填充）。
 */
export function paintPivotUnconfiguredPlaceholderArt(
  ctx: CanvasRenderingContext2D,
  sheet: Worksheet,
  layout: FrozenLayout,
  viewZoom: number,
  scrollX: number,
  scrollY: number,
  headerW: number,
  headerH: number,
  canvasW: number,
  canvasH: number,
  r0: number,
  r1: number,
  c0: number,
  c1: number,
): void {
  for (const def of sheet.getPivotTableDefinitionsSnapshot()) {
    if (!isUnconfiguredPivotDefinition(def)) {
      continue;
    }
    const top = pivotLayoutStartRow(def);
    const cStart = def.destinationCol;
    const rows = Math.max(1, def.outputRowCount);
    const cols = Math.max(1, def.outputColCount);
    const artRow0 = top + 2;
    const artRow1 = top + rows - 1;
    if (artRow0 > artRow1) {
      continue;
    }
    const cEnd = cStart + cols - 1;
    if (artRow1 < r0 || artRow0 > r1 || cEnd < c0 || cStart > c1) {
      continue;
    }

    let x = cellLeftX(sheet, layout, cStart, viewZoom, scrollX);
    let y = cellTopY(sheet, layout, artRow0, viewZoom, scrollY);
    let w = 0;
    let h = 0;
    for (let c = cStart; c <= cEnd; c++) {
      w += scaledColWidthAt(sheet, c, viewZoom);
    }
    for (let r = artRow0; r <= artRow1; r++) {
      h += scaledRowHeightAt(sheet, r, viewZoom);
    }
    if (w < 48 || h < 48) {
      continue;
    }
    if (!cellIntersectsCanvas(x, y, w, h, headerW, headerH, canvasW, canvasH)) {
      continue;
    }

    drawPivotEmptyStateIllustration(ctx, x, y, w, h);
  }
}

function drawPivotEmptyStateIllustration(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
): void {
  ctx.save();
  ctx.beginPath();
  ctx.rect(x, y, w, h);
  ctx.clip();

  ctx.fillStyle = "#FAFBFD";
  ctx.fillRect(x, y, w, h);

  const pad = Math.min(w, h) * 0.06;
  const ix = x + pad;
  const iy = y + pad;
  const iw = w - pad * 2;
  const ih = h - pad * 2;
  if (iw < 32 || ih < 32) {
    ctx.restore();
    return;
  }

  const u = iw / 100;
  const v = ih / 100;
  const line = Math.max(0.8, u * 0.12);

  // 左侧小表
  const tw = 26 * u;
  const th = 50 * v;
  const tx = ix + 2 * u;
  const ty = iy + ih - th - 5 * v;
  ctx.fillStyle = "#8FAADC";
  ctx.fillRect(tx, ty, tw, th * 0.2);
  ctx.fillStyle = "#F2F4F8";
  ctx.fillRect(tx, ty + th * 0.2, tw, th * 0.8);
  ctx.strokeStyle = "#C5CEDD";
  ctx.lineWidth = line;
  for (let i = 1; i <= 4; i++) {
    const ly = ty + th * 0.2 + ((th * 0.8) * i) / 5;
    ctx.beginPath();
    ctx.moveTo(tx, ly);
    ctx.lineTo(tx + tw, ly);
    ctx.stroke();
  }
  ctx.beginPath();
  ctx.moveTo(tx + tw * 0.38, ty + th * 0.2);
  ctx.lineTo(tx + tw * 0.38, ty + th);
  ctx.stroke();

  // 右侧面板（字段区示意）
  const pw = 30 * u;
  const ph = 52 * v;
  const px = ix + iw - pw - 4 * u;
  const py = iy + ih - ph - 4 * v;
  ctx.fillStyle = "#D6E3F4";
  ctx.strokeStyle = "#8FAADC";
  ctx.lineWidth = line;
  ctx.fillRect(px, py, pw, ph);
  ctx.strokeRect(px, py, pw, ph);
  ctx.fillStyle = "#B4C6E7";
  ctx.fillRect(px + line, py + line, pw - line * 2, ph * 0.14);
  const zoneH = (ph * 0.72) / 4;
  const zy0 = py + ph * 0.18;
  for (let zi = 0; zi < 4; zi++) {
    const zy = zy0 + zi * zoneH;
    ctx.fillStyle = zi % 2 === 0 ? "#E8EEF7" : "#DEE8F4";
    ctx.fillRect(px + 3 * u, zy + v * 0.35, pw - 6 * u, zoneH - v * 0.7);
  }

  // 中间放大镜 + 勾选列表
  const cx = ix + iw * 0.48;
  const cy = iy + ih * 0.5;
  const r = Math.min(iw, ih) * 0.19;
  ctx.fillStyle = "rgba(255,255,255,0.96)";
  ctx.strokeStyle = "#6B8EC6";
  ctx.lineWidth = Math.max(1.2, u * 0.22);
  ctx.beginPath();
  ctx.arc(cx, cy, r, 0, Math.PI * 2);
  ctx.fill();
  ctx.stroke();
  const hx = Math.cos(Math.PI / 4) * r;
  const hy = Math.sin(Math.PI / 4) * r;
  ctx.beginPath();
  ctx.moveTo(cx + hx, cy + hy);
  ctx.lineTo(cx + hx + 10 * u, cy + hy + 10 * v);
  ctx.strokeStyle = "#6B8EC6";
  ctx.lineWidth = Math.max(1.5, u * 0.28);
  ctx.stroke();

  const bx = cx - r * 0.55;
  const by = cy - r * 0.42;
  const bw = r * 1.1;
  const bh = r * 0.88;
  const rowGap = bh / 3.2;
  for (let i = 0; i < 3; i++) {
    const ry = by + i * rowGap + rowGap * 0.15;
    const s = Math.min(rowGap * 0.65, r * 0.22);
    ctx.strokeStyle = "#9AA5B5";
    ctx.lineWidth = line;
    ctx.strokeRect(bx, ry, s, s);
    if (i === 1) {
      ctx.strokeStyle = "#E97132";
      ctx.lineWidth = Math.max(1.5, line * 1.4);
      ctx.beginPath();
      ctx.moveTo(bx + s * 0.18, ry + s * 0.52);
      ctx.lineTo(bx + s * 0.42, ry + s * 0.78);
      ctx.lineTo(bx + s * 0.92, ry + s * 0.22);
      ctx.stroke();
    }
  }

  ctx.restore();
}
