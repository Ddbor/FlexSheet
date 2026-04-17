/** 对话框内「在工作表中选定区域」按钮使用的网格图标（currentColor）。 */
export function createRangePickerIconSvg(): SVGSVGElement {
  const ns = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(ns, "svg");
  svg.setAttribute("viewBox", "0 0 24 24");
  svg.setAttribute("fill", "none");
  svg.setAttribute("stroke", "currentColor");
  svg.setAttribute("stroke-width", "1.6");
  const g = document.createElementNS(ns, "g");
  const r = document.createElementNS(ns, "rect");
  r.setAttribute("x", "3");
  r.setAttribute("y", "4");
  r.setAttribute("width", "14");
  r.setAttribute("height", "14");
  r.setAttribute("rx", "1.5");
  const p = document.createElementNS(ns, "path");
  p.setAttribute("d", "M17 8l4-2v12l-4-2M7 9h5M7 12h5M7 15h4");
  g.appendChild(r);
  g.appendChild(p);
  svg.appendChild(g);
  return svg;
}
