import type { CellFillPatternType } from "@flexsheet/core";

/**
 * 在已绘制背景色的单元格矩形上叠加填充图案（前景色）。
 * `none` 不绘制；其余类型与 OOXML / Excel 图案语义近似。
 */
export function paintCellFillPatternOverlay(
  ctx: CanvasRenderingContext2D,
  x: number,
  y: number,
  w: number,
  h: number,
  patternType: CellFillPatternType,
  fgCss: string,
): void {
  if (patternType === "none" || w <= 0 || h <= 0) {
    return;
  }
  ctx.save();
  ctx.beginPath();
  ctx.rect(x, y, w, h);
  ctx.clip();
  ctx.fillStyle = fgCss;
  ctx.strokeStyle = fgCss;
  ctx.lineCap = "square";

  const drawDots = (step: number, alpha: number): void => {
    ctx.globalAlpha = alpha;
    for (let py = y; py < y + h; py += step) {
      for (let px = x; px < x + w; px += step) {
        ctx.fillRect(px, py, 1, 1);
      }
    }
    ctx.globalAlpha = 1;
  };

  const line = (x1: number, y1: number, x2: number, y2: number, lw: number): void => {
    ctx.lineWidth = lw;
    ctx.beginPath();
    ctx.moveTo(x1, y1);
    ctx.lineTo(x2, y2);
    ctx.stroke();
  };

  switch (patternType) {
    case "gray125":
      ctx.globalAlpha = 0.125;
      ctx.fillRect(x, y, w, h);
      ctx.globalAlpha = 1;
      break;
    case "gray0625":
      ctx.globalAlpha = 0.0625;
      ctx.fillRect(x, y, w, h);
      ctx.globalAlpha = 1;
      break;
    /** 约 25% / 50% / 75% 灰点（与 Excel 填充图案第 1 行后三格对应） */
    case "darkGray":
      drawDots(2, 0.78);
      break;
    case "mediumGray":
      drawDots(3, 0.52);
      break;
    case "lightGray":
      drawDots(4, 0.28);
      break;
    case "darkHorizontal":
      ctx.lineWidth = 2;
      for (let py = y; py <= y + h; py += 6) {
        line(x, py, x + w, py, 2);
      }
      break;
    case "lightHorizontal":
      ctx.lineWidth = 1;
      for (let py = y + 1; py <= y + h; py += 7) {
        line(x, py, x + w, py, 1);
      }
      break;
    case "darkVertical":
      ctx.lineWidth = 2;
      for (let px = x; px <= x + w; px += 6) {
        line(px, y, px, y + h, 2);
      }
      break;
    case "lightVertical":
      ctx.lineWidth = 1;
      for (let px = x + 1; px <= x + w; px += 7) {
        line(px, y, px, y + h, 1);
      }
      break;
    /** 斜率 +1（y=x+c）：c 取遍矩形四角上 y−x，才能铺满区域。 */
    case "darkDown": {
      ctx.lineWidth = 2;
      const step = 6;
      for (let c = y - x - w; c <= y - x + h; c += step) {
        const xStart = Math.max(x, y - c);
        const xEnd = Math.min(x + w, y + h - c);
        if (xStart < xEnd) {
          line(xStart, xStart + c, xEnd, xEnd + c, 2);
        }
      }
      break;
    }
    case "lightDown": {
      ctx.lineWidth = 1;
      const step = 8;
      for (let c = y - x - w; c <= y - x + h; c += step) {
        const xStart = Math.max(x, y - c);
        const xEnd = Math.min(x + w, y + h - c);
        if (xStart < xEnd) {
          line(xStart, xStart + c, xEnd, xEnd + c, 1);
        }
      }
      break;
    }
    /** 斜率 −1（x+y=b）：b 取遍四角上 x+y，才能铺满区域。 */
    case "darkUp": {
      ctx.lineWidth = 2;
      const step = 6;
      for (let b = x + y; b <= x + w + y + h; b += step) {
        const xStart = Math.max(x, b - y - h);
        const xEnd = Math.min(x + w, b - y);
        if (xStart < xEnd) {
          line(xStart, -xStart + b, xEnd, -xEnd + b, 2);
        }
      }
      break;
    }
    case "lightUp": {
      ctx.lineWidth = 1;
      const step = 8;
      for (let b = x + y; b <= x + w + y + h; b += step) {
        const xStart = Math.max(x, b - y - h);
        const xEnd = Math.min(x + w, b - y);
        if (xStart < xEnd) {
          line(xStart, -xStart + b, xEnd, -xEnd + b, 1);
        }
      }
      break;
    }
    case "darkGrid": {
      ctx.lineWidth = 2;
      for (let py = y; py <= y + h; py += 6) {
        line(x, py, x + w, py, 2);
      }
      for (let px = x; px <= x + w; px += 6) {
        line(px, y, px, y + h, 2);
      }
      break;
    }
    case "lightGrid": {
      ctx.lineWidth = 1;
      for (let py = y + 1; py <= y + h; py += 7) {
        line(x, py, x + w, py, 1);
      }
      for (let px = x + 1; px <= x + w; px += 7) {
        line(px, y, px, y + h, 1);
      }
      break;
    }
    case "darkTrellis": {
      ctx.lineWidth = 2;
      const step = 6;
      for (let t = x - h; t < x + w + h; t += step) {
        const xA = Math.max(x, t);
        const xB = Math.min(x + w, t + h);
        if (xA < xB) {
          line(xA, y + (xA - t), xB, y + (xB - t), 2);
        }
      }
      for (let t = x - h; t < x + w + h; t += step) {
        const xA = Math.max(x, t);
        const xB = Math.min(x + w, t + h);
        if (xA < xB) {
          line(xA, y + h - (xA - t), xB, y + h - (xB - t), 2);
        }
      }
      break;
    }
    case "lightTrellis": {
      ctx.lineWidth = 1;
      const step = 8;
      for (let t = x - h; t < x + w + h; t += step) {
        const xA = Math.max(x, t);
        const xB = Math.min(x + w, t + h);
        if (xA < xB) {
          line(xA, y + (xA - t), xB, y + (xB - t), 1);
        }
      }
      for (let t = x - h; t < x + w + h; t += step) {
        const xA = Math.max(x, t);
        const xB = Math.min(x + w, t + h);
        if (xA < xB) {
          line(xA, y + h - (xA - t), xB, y + h - (xB - t), 1);
        }
      }
      break;
    }
    default:
      break;
  }
  ctx.restore();
}
