/** Ribbon「条件格式 → 数据条」二级菜单中的彩色缩略图（SVG）。 */

export type DataBarFlyoutThumbFill = "gradient" | "solid";

/**
 * 绘制迷你表格 + 双色列上的横向数据条（仿 Excel 数据条预览）。
 * @param fillKind 渐变（左实色→右白）或实心
 * @param barColorCss 正值条主色，如 `#638ec6`
 */
export function createDataBarFlyoutThumbnail(
  fillKind: DataBarFlyoutThumbFill,
  barColorCss: string,
): SVGSVGElement {
  const ns = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(ns, "svg");
  /* 较扁的 viewBox，显示尺寸略小，避免二级菜单项过高 */
  svg.setAttribute("viewBox", "0 0 40 32");
  svg.setAttribute("width", "32");
  svg.setAttribute("height", "26");
  svg.setAttribute("aria-hidden", "true");

  const gid = `dbft-${Math.random().toString(36).slice(2, 10)}`;

  if (fillKind === "gradient") {
    const def = document.createElementNS(ns, "defs");
    const lg = document.createElementNS(ns, "linearGradient");
    lg.setAttribute("id", gid);
    lg.setAttribute("x1", "0%");
    lg.setAttribute("y1", "0%");
    lg.setAttribute("x2", "100%");
    lg.setAttribute("y2", "0%");
    const s0 = document.createElementNS(ns, "stop");
    s0.setAttribute("offset", "0%");
    s0.setAttribute("stop-color", barColorCss);
    const s1 = document.createElementNS(ns, "stop");
    s1.setAttribute("offset", "100%");
    s1.setAttribute("stop-color", "#ffffff");
    lg.appendChild(s0);
    lg.appendChild(s1);
    def.appendChild(lg);
    svg.appendChild(def);
  }

  const bg = document.createElementNS(ns, "rect");
  bg.setAttribute("width", "40");
  bg.setAttribute("height", "32");
  bg.setAttribute("fill", "#e8e8e8");
  bg.setAttribute("rx", "1.5");
  svg.appendChild(bg);

  const cellW = 17;
  const cellH = 6;
  const gap = 1;
  const ox = 1;
  const oy = 1;
  for (let c = 0; c < 2; c++) {
    for (let r = 0; r < 4; r++) {
      const cell = document.createElementNS(ns, "rect");
      cell.setAttribute("x", String(ox + c * (cellW + gap)));
      cell.setAttribute("y", String(oy + r * (cellH + gap)));
      cell.setAttribute("width", String(cellW));
      cell.setAttribute("height", String(cellH));
      cell.setAttribute("fill", "#f7f7f7");
      cell.setAttribute("stroke", "#d0d0d0");
      cell.setAttribute("stroke-width", "0.35");
      svg.appendChild(cell);
    }
  }

  const col0Frac = [0.78, 0.48, 0.92, 0.36];
  const col1Frac = [0.58, 0.85, 0.42, 0.7];
  for (let c = 0; c < 2; c++) {
    const fr = c === 0 ? col0Frac : col1Frac;
    for (let r = 0; r < 4; r++) {
      const x0 = ox + 1 + c * (cellW + gap);
      const y0 = oy + 1 + r * (cellH + gap) + 2;
      const maxW = cellW - 2;
      const bw = Math.max(1.5, maxW * (fr[r] ?? 0.5));
      const bar = document.createElementNS(ns, "rect");
      bar.setAttribute("x", String(x0));
      bar.setAttribute("y", String(y0));
      bar.setAttribute("width", String(bw));
      bar.setAttribute("height", "3");
      bar.setAttribute("rx", "0.5");
      if (fillKind === "gradient") {
        bar.setAttribute("fill", `url(#${gid})`);
      } else {
        bar.setAttribute("fill", barColorCss);
      }
      svg.appendChild(bar);
    }
  }

  return svg;
}
