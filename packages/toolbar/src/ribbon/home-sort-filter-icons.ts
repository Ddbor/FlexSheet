/**
 * 「开始 → 排序和筛选」专用图标（仅用于该 Ribbon 区与下拉），单色 currentColor。
 */

function svgElMarkup(viewBox: string, innerMarkup: string): SVGSVGElement {
  const ns = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(ns, "svg");
  svg.setAttribute("viewBox", viewBox);
  svg.setAttribute("aria-hidden", "true");
  svg.innerHTML = innerMarkup;
  return svg;
}

/** Ribbon 主钮 32×32：漏斗 + 排序示意 */
export function iconSortFilterRibbon(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    "<g fill-rule='nonzero'>" +
      "<path fill='currentColor' d='M23.5 23v-5.358L31 8H13.5l7.5 9.642V25.5z'/>" +
      "<path fill='currentColor' fill-opacity='.55' d='M6.026 3h2.448L13.5 15.5h-2.301l-1.197-3.134H4.498L3.3 15.5H1zm-.884 7.686h4.216l-2.062-5.48h-.074z'/>" +
      "<path fill='currentColor' fill-opacity='.42' d='M2.47 16.75h9.43v1.588l-6.36 7.638h6.71V28h-10v-1.676l6.272-7.55H2.469z'/>" +
      "</g>",
  );
}

function iconMenu16Markup(inner: string): SVGSVGElement {
  return svgElMarkup("0 0 16 16", "<g fill='none' fill-rule='evenodd'>" + inner + "</g>");
}

/** 升序 */
export function iconMenuSortAsc(): SVGSVGElement {
  return iconMenu16Markup(
    "<path fill='currentColor' fill-rule='nonzero' d='M12 14.994 15 12h-2V1h-2v11H9z'/>" +
      "<path fill='currentColor' fill-opacity='.5' d='M1.806 9h5.45v.782L3.059 14.16H7.5V15h-6v-.773L5.707 9.84H1.806z'/>" +
      "<path fill='currentColor' fill-opacity='.72' d='M3.814 1h1.372L8 8H6.711l-.67-1.755H2.96L2.289 8H1zM3.32 5.304h2.36L4.526 2.235h-.041z'/>",
  );
}

/** 降序 */
export function iconMenuSortDesc(): SVGSVGElement {
  return iconMenu16Markup(
    "<path fill='currentColor' fill-rule='nonzero' d='M12 14.994 15 12h-2V1h-2v11H9z'/>" +
      "<path fill='currentColor' fill-opacity='.5' d='M1.806 1h5.45v.782L3.059 6.16H7.5V7h-6v-.773L5.707 1.84H1.806z'/>" +
      "<path fill='currentColor' fill-opacity='.72' d='M3.814 8h1.372L8 15H6.711l-.67-1.755H2.96L2.289 15H1zm-.494 4.304h2.36L4.526 9.235h-.041z'/>",
  );
}

/** 自定义排序 */
export function iconMenuSortCustom(): SVGSVGElement {
  return iconMenu16Markup(
    "<path fill='currentColor' fill-opacity='.12' d='M1 4h14v11H1z'/>" +
      "<path fill='currentColor' d='M16 4v12H0V0h16zm-1 0H1v11h14z'/>" +
      "<path fill='currentColor' fill-rule='nonzero' d='m5 14 2-3H5.5V5h-1v6H3z'/>" +
      "<path fill='currentColor' fill-opacity='.5' fill-rule='nonzero' d='M11 5 9 8h1.5v6h1V8H13z'/>",
  );
}

/** 筛选 */
export function iconMenuFilterOnly(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    "<path fill='currentColor' fill-rule='nonzero' d='M9 13V8.714L15 1H1l6 7.714V15z'/>",
  );
}

/** 清除筛选 */
export function iconMenuFilterClear(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    "<g fill='none' fill-rule='nonzero'>" +
      "<path fill='currentColor' d='M8.193 6.152c-.672-.215-1.262-.123-1.629.34a2.3 2.3 0 0 0-.46 1.148c-.058.649.26 1.2.951 1.446q.177.063.345.133V10.6l-1.6 1.6V7.171L1 1h11.2z'/>" +
      "<path fill='currentColor' fill-opacity='.58' d='M14.983 7.9c-.258.663-1.648 1.3-2.925 2.248.636 1.011 1.338 2.367 2.506 4.852-1.464-1.713-2.54-2.75-3.625-3.892-1.156 1.057-1.704 1.617-2.107 2.366-.458.852-.722 1.81-1.589 1.394-.578-.278-1.099-1.317 0-2.394q1.65-1.617 2.573-2.415c-.84-.681-1.594-1.15-2.573-1.5s-.493-1.397-.24-1.719c.254-.32.778-.32 1.402 0q1.834 1.165 2.833 2.233c1.436-.76 3.996-1.82 3.745-1.172'/>" +
      "</g>",
  );
}

/** 重新应用（与资源同构路径，二色合为 currentColor 分层） */
export function iconMenuFilterReapply(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    "<g fill='none' fill-rule='nonzero'>" +
      "<path fill='currentColor' d='M5.8 11.69V7.17L1 1h11.2L7.4 7.171v2.117z'/>" +
      "<path fill='currentColor' fill-opacity='.58' d='m7.72 9.819 1.119 1.68h-.7c0 .626.22 1.202.583 1.657l.032.04q.075.091.159.175l.008.01a2.7 2.7 0 0 0 .817.554q.02.01.04.018.09.036.18.067l.075.024q.08.025.16.042l.1.023q.084.014.169.026.045.006.089.01.107.01.217.012a2.64 2.64 0 0 0 1.558-.48.42.42 0 1 1 .482.688 3.48 3.48 0 0 1-2.351.616q-.035-.004-.07-.01-.098-.009-.195-.025-.028-.006-.056-.013-.07-.014-.138-.031a4 4 0 0 1-.313-.089 4 4 0 0 1-.226-.083l-.065-.028a4 4 0 0 1-.254-.124q-.011-.007-.024-.012a3.5 3.5 0 0 1-.786-.59l-.019-.02a4 4 0 0 1-.202-.222l-.048-.062q-.066-.08-.128-.166-.008-.011-.013-.024a3.48 3.48 0 0 1-.621-1.984H6.6zM10.8 8l.05.003q.138 0 .278.013l.114.015q.084.008.167.02l.06.015.123.027q.111.025.22.058.045.013.09.03.124.04.246.09l.042.018c.505.218.952.55 1.307.982l.018.023q.08.096.153.198l.016.028c.388.563.616 1.244.616 1.979h.7l-1.12 1.68-1.12-1.68h.7c0-.639-.227-1.226-.605-1.685a2.66 2.66 0 0 0-1.007-.758l-.02-.01a3 3 0 0 0-.199-.072l-.056-.02a3 3 0 0 0-.18-.045q-.04-.01-.082-.019a3 3 0 0 0-.187-.029l-.07-.008a2.641 2.641 0 0 0-1.779.467.419.419 0 1 1-.483-.687A3.5 3.5 0 0 1 10.8 8'/>" +
      "</g>",
  );
}
