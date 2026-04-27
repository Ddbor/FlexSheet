/**
 * 「开始 → 查找」Ribbon 图标，单色 `currentColor`（与 `home-sort-filter-icons` 中工厂一致）。
 */

function svgElMarkup(viewBox: string, innerMarkup: string): SVGSVGElement {
  const ns = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(ns, "svg");
  svg.setAttribute("viewBox", viewBox);
  svg.setAttribute("aria-hidden", "true");
  svg.innerHTML = innerMarkup;
  return svg;
}

/** Ribbon 主钮 32×32：放大镜 */
export function iconFindRibbon(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    "<path fill='currentColor' fill-rule='nonzero' d='M27.695 26.482c.376.376.409.952.074 1.287s-.91.303-1.287-.073l-6.547-6.548a10.25 10.25 0 0 1-6.642 2.437C7.609 23.585 3 18.977 3 13.293S7.609 3 13.293 3s10.292 4.609 10.292 10.293c0 2.533-.92 4.85-2.437 6.642zM13.293 21.87a8.577 8.577 0 0 0 8.577-8.577 8.577 8.577 0 0 0-8.577-8.578 8.577 8.577 0 0 0-8.578 8.578 8.577 8.577 0 0 0 8.578 8.577' />",
  );
}
