/** Ribbon「条件格式 → 色阶」二级菜单中的彩色缩略图（SVG）。 */

export type ColorScaleThumbKind = "two" | "three";

/**
 * 迷你色阶条：双色为左右渐变；三色为左中右三色渐变。
 */
export function createColorScaleFlyoutThumbnail(
  kind: ColorScaleThumbKind,
  minHex: string,
  maxHex: string,
  midHex?: string,
): SVGSVGElement {
  const ns = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(ns, "svg");
  svg.setAttribute("viewBox", "0 0 40 26");
  svg.setAttribute("width", "32");
  svg.setAttribute("height", "22");
  svg.setAttribute("aria-hidden", "true");
  svg.setAttribute("class", "fs-bd-cs-flyout-thumb");

  const gid = `csft-${Math.random().toString(36).slice(2, 10)}`;

  const def = document.createElementNS(ns, "defs");
  const lg = document.createElementNS(ns, "linearGradient");
  lg.setAttribute("id", gid);
  lg.setAttribute("x1", "0%");
  lg.setAttribute("y1", "0%");
  lg.setAttribute("x2", "100%");
  lg.setAttribute("y2", "0%");

  const s0 = document.createElementNS(ns, "stop");
  s0.setAttribute("offset", "0%");
  s0.setAttribute("stop-color", minHex);
  lg.appendChild(s0);

  if (kind === "three" && midHex !== undefined && midHex !== "") {
    const sm = document.createElementNS(ns, "stop");
    sm.setAttribute("offset", "50%");
    sm.setAttribute("stop-color", midHex);
    lg.appendChild(sm);
  }

  const s1 = document.createElementNS(ns, "stop");
  s1.setAttribute("offset", "100%");
  s1.setAttribute("stop-color", maxHex);
  lg.appendChild(s1);

  def.appendChild(lg);
  svg.appendChild(def);

  const frame = document.createElementNS(ns, "rect");
  frame.setAttribute("x", "1");
  frame.setAttribute("y", "1");
  frame.setAttribute("width", "38");
  frame.setAttribute("height", "24");
  frame.setAttribute("rx", "2");
  frame.setAttribute("fill", "#ececec");
  frame.setAttribute("stroke", "#c8c6c4");
  frame.setAttribute("stroke-width", "0.5");
  svg.appendChild(frame);

  const bar = document.createElementNS(ns, "rect");
  bar.setAttribute("x", "3");
  bar.setAttribute("y", "5");
  bar.setAttribute("width", "34");
  bar.setAttribute("height", "16");
  bar.setAttribute("rx", "1.5");
  bar.setAttribute("fill", `url(#${gid})`);
  bar.setAttribute("stroke", "#b0b0b0");
  bar.setAttribute("stroke-width", "0.35");
  svg.appendChild(bar);

  return svg;
}
