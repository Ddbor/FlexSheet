import { resolveCellBorderStroke } from "@flexsheet/core";
import type { CellBorderLinePattern, CellBorderSide } from "@flexsheet/core";

import type { FormatCellsLineSwatchId } from "./format-cells-border.js";

const NS = "http://www.w3.org/2000/svg";

function sideForSwatch(swatch: Exclude<FormatCellsLineSwatchId, "none">): CellBorderSide {
  return { kind: "thin", linePattern: swatch as CellBorderLinePattern };
}

function mkLine(
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  color: string,
  sw: number,
  dash: string | undefined,
  linecap: CanvasLineCap,
): SVGLineElement {
  const el = document.createElementNS(NS, "line");
  el.setAttribute("x1", String(x1));
  el.setAttribute("y1", String(y1));
  el.setAttribute("x2", String(x2));
  el.setAttribute("y2", String(y2));
  el.setAttribute("stroke", color);
  el.setAttribute("stroke-width", String(sw));
  el.setAttribute("stroke-linecap", linecap);
  if (dash !== undefined) {
    el.setAttribute("stroke-dasharray", dash);
  }
  return el;
}

/** 在任意方向线段上绘制与线型网格、Canvas 一致的笔画（数据来自 `resolveCellBorderStroke`）。 */
export function appendFormatCellsSwatchStroke(
  parent: SVGSVGElement | SVGGElement,
  x1: number,
  y1: number,
  x2: number,
  y2: number,
  swatch: Exclude<FormatCellsLineSwatchId, "none">,
  color: string,
): void {
  const r = resolveCellBorderStroke(sideForSwatch(swatch), 1);
  const cap = r.lineCap;

  if (r.double) {
    const len = Math.hypot(x2 - x1, y2 - y1) || 1;
    const nx = (-(y2 - y1) / len) * (r.gap / 2);
    const ny = ((x2 - x1) / len) * (r.gap / 2);
    const w = r.lineWidth;
    parent.appendChild(mkLine(x1 + nx, y1 + ny, x2 + nx, y2 + ny, color, w, undefined, cap));
    parent.appendChild(mkLine(x1 - nx, y1 - ny, x2 - nx, y2 - ny, color, w, undefined, cap));
    return;
  }

  const dashStr =
    r.lineDash !== null && r.lineDash.length > 0 ? r.lineDash.map((n) => String(n)).join(" ") : undefined;
  parent.appendChild(mkLine(x1, y1, x2, y2, color, r.lineWidth, dashStr, cap));
}

/**
 * 线型预览：与 Excel「设置单元格格式」边框选项卡线型网格一致（黑色、水平居中）。
 */
export function createFormatCellsLineSwatchHost(swatch: FormatCellsLineSwatchId): HTMLElement {
  const host = document.createElement("span");
  host.className = "fs-format-cells__border-line-swatch";
  if (swatch === "none") {
    host.textContent = "无";
    host.style.display = "flex";
    host.style.alignItems = "center";
    host.style.justifyContent = "center";
    host.style.fontSize = "12px";
    host.style.fontWeight = "400";
    host.style.color = "#000000";
    host.style.fontFamily = 'system-ui, "Segoe UI", "Microsoft YaHei", sans-serif';
    return host;
  }

  const svg = document.createElementNS(NS, "svg");
  svg.setAttribute("viewBox", "0 0 96 16");
  svg.setAttribute("width", "100%");
  svg.setAttribute("height", "16");
  svg.style.display = "block";
  svg.style.maxWidth = "96px";
  svg.setAttribute("aria-hidden", "true");

  appendFormatCellsSwatchStroke(svg, 4, 8, 92, 8, swatch, "#000000");

  host.appendChild(svg);
  return host;
}

/** 「边框」预览框内叠加 SVG，与当前选中线型一致。 */
export function createFormatCellsBorderPreviewSvg(options: {
  readonly swatch: FormatCellsLineSwatchId;
  readonly colorCss: string;
  readonly top: boolean;
  readonly bottom: boolean;
  readonly left: boolean;
  readonly right: boolean;
  readonly insideH: boolean;
  readonly insideV: boolean;
  readonly multiCell: boolean;
}): SVGSVGElement {
  const svg = document.createElementNS(NS, "svg");
  svg.setAttribute("viewBox", "0 0 132 88");
  svg.setAttribute("preserveAspectRatio", "none");
  svg.style.position = "absolute";
  svg.style.inset = "0";
  svg.style.width = "100%";
  svg.style.height = "100%";
  svg.style.pointerEvents = "none";
  svg.setAttribute("aria-hidden", "true");

  const { swatch, colorCss, multiCell } = options;
  if (swatch === "none") {
    return svg;
  }

  const m = 6;
  const W = 132;
  const H = 88;
  const x0 = m;
  const x1 = W - m;
  const y0 = m;
  const y1 = H - m;
  const midX = W / 2;
  const midY = H / 2;
  const c = colorCss;

  const draw = (ax0: number, ay0: number, ax1: number, ay1: number): void => {
    appendFormatCellsSwatchStroke(svg, ax0, ay0, ax1, ay1, swatch, c);
  };

  if (options.top) {
    draw(x0, y0, x1, y0);
  }
  if (options.bottom) {
    draw(x0, y1, x1, y1);
  }
  if (options.left) {
    draw(x0, y0, x0, y1);
  }
  if (options.right) {
    draw(x1, y0, x1, y1);
  }
  if (multiCell) {
    if (options.insideH) {
      draw(x0, midY, x1, midY);
    }
    if (options.insideV) {
      draw(midX, y0, midX, y1);
    }
  }

  return svg;
}
