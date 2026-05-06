/**
 * 内联 SVG 图标（不依赖 Font Awesome 包，避免新增依赖）。
 * 风格接近 Office 线性图标。
 */

function svgEl(viewBox: string, paths: string, ariaHidden = true): SVGSVGElement {
  const ns = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(ns, "svg");
  svg.setAttribute("viewBox", viewBox);
  svg.setAttribute("fill", "none");
  svg.setAttribute("stroke", "currentColor");
  svg.setAttribute("stroke-width", "1.5");
  svg.setAttribute("stroke-linecap", "round");
  svg.setAttribute("stroke-linejoin", "round");
  if (ariaHidden) {
    svg.setAttribute("aria-hidden", "true");
  }
  svg.innerHTML = paths;
  return svg;
}

/** 内联填充路径（与资源 viewBox 一致，路径内自管 fill） */
function svgElMarkup(viewBox: string, innerMarkup: string): SVGSVGElement {
  const ns = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(ns, "svg");
  svg.setAttribute("viewBox", viewBox);
  svg.setAttribute("aria-hidden", "true");
  svg.innerHTML = innerMarkup;
  return svg;
}

/** 粘贴（32×32 剪贴板，单色 currentColor，层次用透明度区分） */
export function iconPaste(): SVGSVGElement {
  const ns = "http://www.w3.org/2000/svg";
  const svg = document.createElementNS(ns, "svg");
  svg.setAttribute("viewBox", "0 0 32 32");
  svg.setAttribute("aria-hidden", "true");
  /* 比 Ribbon 主文字略浅，与其它线性图标观感接近 */
  svg.setAttribute("opacity", "0.58");
  svg.innerHTML =
    '<g fill-rule="nonzero">' +
    '<path fill="currentColor" d="M9 3v2a2 2 0 0 0 2 2h10a2 2 0 0 0 2-2V3h3a1 1 0 0 1 1 1v7H12v17H6a1 1 0 0 1-1-1V4a1 1 0 0 1 1-1z"/>' +
    '<path fill="currentColor" d="M13 2V1a1 1 0 0 1 1-1h4a1 1 0 0 1 1 1v1h2a1 1 0 0 1 1 1v2a1 1 0 0 1-1 1H11a1 1 0 0 1-1-1V3a1 1 0 0 1 1-1zm2-1v2h2V1z"/>' +
    '<path fill="currentColor" fill-opacity="0.42" d="M13 12h11.478L30 17.555V32H13z"/>' +
    '<path fill="currentColor" d="M24 13v5h5v1h-6v-6h-9v18h15V17.967L24.062 13zm-11-1h11.478L30 17.555V32H13z"/>' +
    "</g>";
  return svg;
}

export function iconCut(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" d="m9.428 8.027 1.818 2.166-1.533 1.285-1.59-1.896-1.59 1.896L5 10.193l1.817-2.166-3.531-4.21A2 2 0 0 1 3.532 1l4.59 5.471L12.714 1a2 2 0 0 1 .247 2.818z"/>' +
      '<path fill="currentColor" fill-opacity="0.38" fill-rule="nonzero" d="M4 15a2 2 0 1 0 0-4 2 2 0 0 0 0 4m0 1a3 3 0 1 1 0-6 3 3 0 0 1 0 6M12 15a2 2 0 1 0 0-4 2 2 0 0 0 0 4m0 1a3 3 0 1 1 0-6 3 3 0 0 1 0 6"/>' +
      "</g>",
  );
}

export function iconCopy(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-opacity="0.32" d="M9.951 3H4v9H1V0h6.032z"/>' +
      '<path fill="currentColor" fill-rule="nonzero" d="M6 3V1h.61H2v10h2v1H1V0h6.032l2.92 3H8.555L7 1.401V3z"/>' +
      '<path fill="currentColor" fill-opacity="0.48" d="M4 4v1H3V4zM4 7v1H3V7z"/>' +
      '<path fill="currentColor" fill-opacity="0.32" d="M5 4h6.032L15 8.078V16H5z"/>' +
      '<path fill="currentColor" fill-rule="nonzero" d="M14 8.484V9h-4V5h.61H6v10h8zM13.529 8 11 5.401V8zM5 4h6.032L15 8.078V16H5z"/>' +
      '<path fill="currentColor" fill-opacity="0.48" d="M7 8h2v1H7zM7 11h6v1H7z"/>' +
      "</g>",
  );
}

export function iconFormatPainter(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="nonzero">' +
      '<path fill="currentColor" fill-opacity="0.3" d="m6.659 5.148-.176.175C5.267 6.508 4.283 7.21 3.47 7.438c-.557.156-1.345.217-2.382.19l-.376-.015 6.19 7.241 5.077-4.396z"/>' +
      '<path fill="currentColor" fill-opacity="0.72" d="m6.659 4.142 6.3 6.515L6.816 16 0 8.023V6.702c.027 0 .124.078.318.105q.114.016.442 0 1.7.074 2.49-.147.748-.21 3.409-2.518m0 1.006-.176.175C5.267 6.508 4.283 7.21 3.47 7.438c-.557.156-1.345.217-2.382.19l-.376-.015 6.19 7.241 5.077-4.396z"/>' +
      '<path fill="currentColor" fill-opacity="0.42" d="m3.784 11.15 3.106 3.682 2.852-2.623-.166.003a15 15 0 0 1-2.556-.232l-.353-.063a19 19 0 0 1-2.51-.638z"/>' +
      '<path fill="currentColor" fill-opacity="0.72" d="M0 7.919c1.565 1.578 4.48 3.007 6.763 3.404s4.54.625 5.707-.372L6.816 16zm3.784 3.231 3.106 3.682 2.852-2.623-.166.003a15 15 0 0 1-2.556-.232l-.353-.063a19 19 0 0 1-2.51-.638z"/>' +
      '<path fill="currentColor" fill-opacity="0.28" d="M14.962 2c-.25-.233-.582-.269-1.11-.036l-1.189 1.435-.577.744-.342.427a8 8 0 0 1-.722.7q-.282.23-.905-.181c-.435-.43-1.21-1.335-1.6-1.335q-.39 0-1.378.848l5.312 5.52q.885-.783.971-1.155t-1.205-1.634l-.448-.517a1.24 1.24 0 0 1-.126-.591q0-.357.39-.725l.33-.284.604-.497 1.421-1.135.695-.625c.168-.41.137-.674-.057-.894z"/>' +
      '<path fill="currentColor" d="M15.517 1.411q.879.822.159 2.188l-1.46 1.154-.782.629-.73.606-.183.161-.105.103q-.032.038-.026.044l.37.425.33.39.287.354q.87 1.097.87 1.46c0 .474-.705 1.215-1.83 2.224l-6.41-6.508 1.077-.94c.714-.513 1.063-.71 1.375-.71.39 0 1.262.571 2.227 1.523q.01.009.047-.023l.103-.103.16-.184.302-.37.52-.664.832-1.092.502-.667q1.485-.822 2.365 0m-.555.59c-.25-.234-.582-.27-1.11-.037l-1.189 1.435-.577.744-.342.427a8 8 0 0 1-.722.7q-.282.23-.905-.181c-.435-.43-1.21-1.335-1.6-1.335q-.39 0-1.378.848l5.312 5.52q.885-.783.971-1.155t-1.205-1.634l-.448-.517a1.24 1.24 0 0 1-.126-.591q0-.357.39-.725l.33-.284.604-.497 1.421-1.135.695-.625c.168-.41.137-.674-.057-.894z"/>' +
      "</g>",
  );
}

/** 增大字号（16×16，单色 currentColor） */
export function iconFontGrow(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="nonzero">' +
      '<path fill="currentColor" d="M4.727 9.973h4.476L7.19 4.568q-.096-.261-.207-.906h-.044a5 5 0 0 1-.215.906zM7.921 2 13 15h-1.903l-1.36-3.559H4.185L2.902 15H1L6.079 2z"/>' +
      '<path fill="currentColor" fill-opacity="0.5" d="M13.5 2 16 5h-5z"/>' +
      "</g>",
  );
}

/** 减小字号（16×16，单色 currentColor） */
export function iconFontShrink(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="nonzero">' +
      '<path fill="currentColor" d="M5 9.714h3l-1.35-3.67a4 4 0 0 1-.138-.615h-.03q-.063.397-.144.615zM7.19 4 11 14H9.573l-1.02-2.737H4.389L3.427 14H2L5.81 4z"/>' +
      '<path fill="currentColor" fill-opacity="0.5" d="M11.5 5 9 2h5z"/>' +
      "</g>",
  );
}

export function iconBold(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<path fill="currentColor" fill-rule="nonzero" d="M4 14V1h4.346q2 0 3.082.806 1.082.805 1.082 2.232 0 1.057-.672 1.855a3.44 3.44 0 0 1-1.69 1.1v.033q1.28.168 2.066 1.007.786.84.786 2.106 0 1.772-1.185 2.816Q10.63 14 8.575 14zM6.608 3.065v3.214h1.24q.893 0 1.407-.453.513-.453.513-1.268 0-1.493-2.054-1.493zm0 5.304v3.566h1.548q.988 0 1.545-.486.557-.487.557-1.352 0-.822-.553-1.275t-1.572-.453z"/>',
  );
}

export function iconItalic(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<path fill="currentColor" fill-rule="nonzero" d="M12 2h-1.5l-3 11H9v1H5v-1h1.5l3-11H8V1h4z"/>',
  );
}

export function iconUnderline(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="currentColor" fill-rule="nonzero">' +
      '<path d="M4 13h8v1H4zM12 7.411q0 4.59-4.086 4.589Q4 12 4 7.572V1h1.015v6.51c0 2.164 1.247 3.532 3.03 3.532 1.723 0 2.96-1.333 2.96-3.428V1H12z"/>' +
      "</g>",
  );
}

/** 双下划线（16×16，单色） */
export function iconDoubleUnderline(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="currentColor" fill-rule="nonzero">' +
      '<path d="M4 14h8v1H4zM4 12h8v1H4zM4 11V1h2.804q2.233 0 3.715 1.323Q11.999 3.647 12 5.874q0 2.324-1.481 3.725T6.703 11zm1.03-8.979V10h1.666c1.248 0 2.528-.488 3.228-1.181q1.05-1.039 1.05-2.912T9.896 3.092c-.72-.629-1.971-1.07-3.143-1.07z"/>' +
      "</g>",
  );
}

/** 单元格下边框（16×16，单色） */
export function iconBorderBottom(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="nonzero">' +
      '<path fill="currentColor" d="M2 13h13v1H2z"/>' +
      '<path fill="currentColor" fill-opacity="0.5" d="M2 9h1v1H2zm0 2h1v1H2zm0-4h1v1H2zm0-2h1v1H2zm0-2h1v1H2zm0-2h1v1H2zm2 0h1v1H4zm0 6h1v1H4zm2-6h1v1H6zm0 6h1v1H6zm2-6h1v1H8zm0 2h1v1H8zm0 2h1v1H8zm0 2h1v1H8zm6-4h1v1h-1zm0 4h1v1h-1zm0 2h1v1h-1zm0 2h1v1h-1zM8 9h1v1H8zm0 2h1v1H8zm2-10h1v1h-1zm0 6h1v1h-1zm2-6h1v1h-1zm0 6h1v1h-1zm2-6h1v1h-1zm0 4h1v1h-1z"/>' +
      "</g>",
  );
}

/** 单元格上边框（16×16，单色） */
export function iconBorderTop(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" d="M2 1h13v1H2z"/>' +
      '<path fill="currentColor" fill-opacity="0.5" d="M2 9h1v1H2zm0 2h1v1H2zm0-4h1v1H2zm0-2h1v1H2zm0-2h1v1H2zm0 10h1v1H2zm2 0h1v1H4zm0-6h1v1H4zm2 6h1v1H6zm0-6h1v1H6zm2 6h1v1H8zM8 3h1v1H8zm0 2h1v1H8zm0 2h1v1H8zm6-4h1v1h-1zm0 4h1v1h-1zm0 2h1v1h-1zm0 2h1v1h-1zM8 9h1v1H8zm0 2h1v1H8zm2 2h1v1h-1zm0-6h1v1h-1zm2 6h1v1h-1zm0-6h1v1h-1zm2 6h1v1h-1zm0-8h1v1h-1z"/>' +
      "</g>",
  );
}

/** 单元格左边框（16×16，单色） */
export function iconBorderLeft(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" d="M2 1h1v13H2z"/>' +
      '<path fill="currentColor" fill-opacity="0.5" d="M10 1h1v1h-1zm2 0h1v1h-1zM8 1h1v1H8zM6 1h1v1H6zM4 1h1v1H4zm10 0h1v1h-1zM4 13h1v1H4zm0-6h1v1H4zm2 6h1v1H6zm0-6h1v1H6zm2 6h1v1H8zM8 3h1v1H8zm0 2h1v1H8zm0 2h1v1H8zm6-4h1v1h-1zm0 4h1v1h-1zm0 2h1v1h-1zm0 2h1v1h-1zM8 9h1v1H8zm0 2h1v1H8zm2 2h1v1h-1zm0-6h1v1h-1zm2 6h1v1h-1zm0-6h1v1h-1zm2 6h1v1h-1zm0-8h1v1h-1z"/>' +
      "</g>",
  );
}

/** 单元格右边框（16×16，单色） */
export function iconBorderRight(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" d="M14 1h1v13h-1z"/>' +
      '<path fill="currentColor" fill-opacity="0.5" d="M10 1h1v1h-1zm2 0h1v1h-1zM8 1h1v1H8zM6 1h1v1H6zM4 1h1v1H4zM2 1h1v1H2zm2 12h1v1H4zm0-6h1v1H4zm2 6h1v1H6zm0-6h1v1H6zm2 6h1v1H8zM8 3h1v1H8zm0 2h1v1H8zm0 2h1v1H8zM2 3h1v1H2zm0 4h1v1H2zm0 2h1v1H2zm0 2h1v1H2zm6-2h1v1H8zm0 2h1v1H8zm2 2h1v1h-1zm0-6h1v1h-1zm2 6h1v1h-1zm0-6h1v1h-1zM2 13h1v1H2zm0-8h1v1H2z"/>' +
      "</g>",
  );
}

/** 无边框（仅网格点，16×16，单色） */
export function iconBorderNone(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<path fill="currentColor" fill-rule="evenodd" d="M10 1h1v1h-1zm2 0h1v1h-1zM8 1h1v1H8zM6 1h1v1H6zM4 1h1v1H4zM2 1h1v1H2zm12 0h1v1h-1zM4 13h1v1H4zm0-6h1v1H4zm2 6h1v1H6zm0-6h1v1H6zm2 6h1v1H8zM8 3h1v1H8zm0 2h1v1H8zm0 2h1v1H8zM2 3h1v1H2zm12 0h1v1h-1zM2 7h1v1H2zm12 0h1v1h-1zM2 9h1v1H2zm12 0h1v1h-1zM2 11h1v1H2zm12 0h1v1h-1zM8 9h1v1H8zm0 2h1v1H8zm2 2h1v1h-1zm0-6h1v1h-1zm2 6h1v1h-1zm0-6h1v1h-1zM2 13h1v1H2zm12 0h1v1h-1zM2 5h1v1H2zm12 0h1v1h-1z"/>',
  );
}

/** 所有框线（16×16，单色） */
export function iconBorderAll(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<path fill="currentColor" fill-rule="evenodd" d="M14 8H9v5h5zm0-1V2H9v5zM8 8H3v5h5zm0-1V2H3v5zm6 7H2V1h13v13z"/>',
  );
}

/** 外侧边框（16×16，单色） */
export function iconBorderOutside(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" d="M14 14H2V1h13v13zm0-1V2H3v11z"/>' +
      '<path fill="currentColor" fill-opacity="0.5" d="M4 7h1v1H4zm2 0h1v1H6zm2-4h1v1H8zm0 2h1v1H8zm0 2h1v1H8zm0 2h1v1H8zm0 2h1v1H8zm2-4h1v1h-1zm2 0h1v1h-1z"/>' +
      "</g>",
  );
}

/** 粗匣框线（16×16，单色） */
export function iconBorderThickBox(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" d="M13 14H2V1h13v13zm0-2V3H4v9z"/>' +
      '<path fill="currentColor" fill-opacity="0.5" d="M4 7h1v1H4zm2 0h1v1H6zm2-4h1v1H8zm0 2h1v1H8zm0 2h1v1H8zm0 2h1v1H8zm0 2h1v1H8zm2-4h1v1h-1zm2 0h1v1h-1z"/>' +
      "</g>",
  );
}

/** 双底框线（16×16，单色） */
export function iconBorderDoubleBottom(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" d="M2 13h13v1H2zM2 11h13v1H2z"/>' +
      '<path fill="currentColor" fill-opacity="0.5" d="M2 9h1v1H2zm0-2h1v1H2zm0-2h1v1H2zm0-2h1v1H2zm0-2h1v1H2zm2 0h1v1H4zm0 6h1v1H4zm2-6h1v1H6zm0 6h1v1H6zm2-6h1v1H8zm0 2h1v1H8zm0 2h1v1H8zm0 2h1v1H8zm6-4h1v1h-1zm0 4h1v1h-1zm0 2h1v1h-1zM8 9h1v1H8zm2-8h1v1h-1zm0 6h1v1h-1zm2-6h1v1h-1zm0 6h1v1h-1zm2-6h1v1h-1zm0 4h1v1h-1z"/>' +
      "</g>",
  );
}

/** 粗底框线（16×16，单色） */
export function iconBorderThickBottom(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" d="M2 12h13v2H2z"/>' +
      '<path fill="currentColor" fill-opacity="0.5" d="M2 9h1v1H2zm0 2h1v1H2zm0-4h1v1H2zm0-2h1v1H2zm0-2h1v1H2zm0-2h1v1H2zm2 0h1v1H4zm0 6h1v1H4zm2-6h1v1H6zm0 6h1v1H6zm2-6h1v1H8zm0 2h1v1H8zm0 2h1v1H8zm0 2h1v1H8zm6-4h1v1h-1zm0 4h1v1h-1zm0 2h1v1h-1zm0 2h1v1h-1zM8 9h1v1H8zm0 2h1v1H8zm2-10h1v1h-1zm0 6h1v1h-1zm2-6h1v1h-1zm0 6h1v1h-1zm2-6h1v1h-1zm0 4h1v1h-1z"/>' +
      "</g>",
  );
}

/** 上下框线（16×16，单色） */
export function iconBorderTopBottom(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" d="M2 13h13v1H2zM2 1h13v1H2z"/>' +
      '<path fill="currentColor" fill-opacity="0.5" d="M2 9h1v1H2zm0 2h1v1H2zm0-4h1v1H2zm0-2h1v1H2zm0-2h1v1H2zm2 4h1v1H4zm2 0h1v1H6zm2-4h1v1H8zm0 2h1v1H8zm0 2h1v1H8zm6-4h1v1h-1zm0 4h1v1h-1zm0 2h1v1h-1zm0 2h1v1h-1zM8 9h1v1H8zm0 2h1v1H8zm2-4h1v1h-1zm2 0h1v1h-1zm2-2h1v1h-1z"/>' +
      "</g>",
  );
}

/** 上框线 + 粗下框线（16×16，单色） */
export function iconBorderTopThickBottom(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" d="M2 12h13v2H2zM2 1h13v1H2z"/>' +
      '<path fill="currentColor" fill-opacity="0.5" d="M2 9h1v1H2zm0 2h1v1H2zm0-4h1v1H2zm0-2h1v1H2zm0-2h1v1H2zm2 4h1v1H4zm2 0h1v1H6zm2-4h1v1H8zm0 2h1v1H8zm0 2h1v1H8zm6-4h1v1h-1zm0 4h1v1h-1zm0 2h1v1h-1zm0 2h1v1h-1zM8 9h1v1H8zm0 2h1v1H8zm2-4h1v1h-1zm2 0h1v1h-1zm2-2h1v1h-1z"/>' +
      "</g>",
  );
}

/** 上框线 + 双下框线（16×16，单色） */
export function iconBorderTopDoubleBottom(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" d="M2 13h13v1H2zM2 11h13v1H2zM2 1h13v1H2z"/>' +
      '<path fill="currentColor" fill-opacity="0.5" d="M2 9h1v1H2zm0-2h1v1H2zm0-2h1v1H2zm0-2h1v1H2zm2 4h1v1H4zm2 0h1v1H6zm2-4h1v1H8zm0 2h1v1H8zm0 2h1v1H8zm6-4h1v1h-1zm0 4h1v1h-1zm0 2h1v1h-1zM8 9h1v1H8zm2-2h1v1h-1zm2 0h1v1h-1zm2-2h1v1h-1z"/>' +
      "</g>",
  );
}

/** 填充颜色（16×16，单色层次） */
export function iconFillColor(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-opacity="0.08" d="M1 4h14v11H1z"/>' +
      '<path fill="currentColor" fill-opacity="0.82" d="M16 4v12H0V0h16zm-1 0H1v11h14z"/>' +
      '<path fill="currentColor" fill-rule="nonzero" d="m8 14 4-4H9V5H7v5H4z"/>' +
      "</g>",
  );
}

/** 字体填充颜色（16×16，仿 Office 油漆桶，非彩色） */
export function iconFontFillColor(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="nonzero">' +
      '<path fill="currentColor" fill-opacity="0.16" d="M0 12h16v4H0z"/>' +
      '<path fill="currentColor" fill-opacity="0.72" d="M1 13v2h14v-2zm-1-1h16v4H0z"/>' +
      '<path fill="currentColor" fill-opacity="0.2" d="M6.993 3.33 2.331 7.994l3.676 3.676 4.662-4.662z"/>' +
      '<path fill="currentColor" d="M6.644 3.68 2.33 7.994l3.676 3.676 4.662-4.662L7.592 3.93v2.525l.006.044L8.1 7l-1 1.002L6.1 7l.544-.544zm0-1.33v-.84c0-.323-.204-.547-.493-.553H6.06c-.28 0-.511.24-.511.553v1.935zM4.6 4.393V1.51c0-.324.1-.636.282-.893q.145-.205.346-.35C5.47.094 5.758 0 6.06 0h.073c.255 0 .5.067.715.192.458.266.744.769.744 1.318V2.6L12 7.007 6.007 13 1 7.993z"/>' +
      '<path fill="currentColor" fill-opacity="0.82" d="M11.739 9.267c0 1.181.248 2.142.638 2.733 0-.628.14-1.035.53-1.773.922-1.514 1.594-2.733.603-3.988Q12.518 4.983 10 5s1.393 1.595 1.739 1.977c.142.628 0 1.108 0 2.29"/>' +
      "</g>",
  );
}

/** 向下填充（16×16，单色） */
export function iconFillDown(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-opacity="0.26" d="M1 1h14v14H1z"/>' +
      '<path fill="currentColor" d="M2 2v12h12V2zm13-1v14H1V1z"/>' +
      '<path fill="currentColor" fill-rule="nonzero" d="m8 13 3-3H9V4H7v6H5z"/>' +
      "</g>",
  );
}

/** 向右填充（16×16，单色） */
export function iconFillRight(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-opacity="0.26" d="M1 1h14v14H1z"/>' +
      '<path fill="currentColor" d="M2 2v12h12V2zm13-1v14H1V1z"/>' +
      '<path fill="currentColor" fill-rule="nonzero" d="m13 8-3-3v2H4v2h6v2z"/>' +
      "</g>",
  );
}

/** 向上填充（16×16，单色） */
export function iconFillUp(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-opacity="0.26" d="M1 1h14v14H1z"/>' +
      '<path fill="currentColor" d="M2 2v12h12V2zm13-1v14H1V1z"/>' +
      '<path fill="currentColor" fill-rule="nonzero" d="m8 3-3 3h2v6h2V6h2z"/>' +
      "</g>",
  );
}

/** 向左填充（16×16，单色） */
export function iconFillLeft(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-opacity="0.26" d="M1 1h14v14H1z"/>' +
      '<path fill="currentColor" d="M2 2v12h12V2zm13-1v14H1V1z"/>' +
      '<path fill="currentColor" fill-rule="nonzero" d="m3 8 3 3V9h6V7H6V5z"/>' +
      "</g>",
  );
}

/** 字体颜色（16×16，单色层次） */
export function iconFontColor(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="nonzero">' +
      '<path fill="currentColor" fill-opacity="0.4" d="M0 12h16v4H0z"/>' +
      '<path fill="currentColor" d="M6.106 6.746h3.73L8.158 2.173a5 5 0 0 1-.172-.767h-.037a4.2 4.2 0 0 1-.179.767zM8.768 0 13 11h-1.586l-1.133-3.011H5.654L4.585 11H3L7.232 0z"/>' +
      "</g>",
  );
}

/** 水平左对齐（16×16，与资源几何一致，单色） */
export function iconAlignLeft(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="currentColor" fill-rule="evenodd">' +
      '<path d="M1 0h14v1H1zM1 6h14v1H1zM1 12h14v1H1zM1 3h10v1H1zM1 15h10v1H1zM1 9h10v1H1z"/>' +
      "</g>",
  );
}

/** 水平居中（16×16） */
export function iconAlignCenter(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="currentColor" fill-rule="evenodd">' +
      '<path d="M1 0h14v1H1zM1 6h14v1H1zM1 12h14v1H1zM3 3h10v1H3zM3 15h10v1H3zM3 9h10v1H3z"/>' +
      "</g>",
  );
}

/** 水平右对齐（16×16） */
export function iconAlignRight(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="currentColor" fill-rule="evenodd">' +
      '<path d="M1 0h14v1H1zM1 6h14v1H1zM1 12h14v1H1zM5 3h10v1H5zM5 9h10v1H5zM5 15h10v1H5z"/>' +
      "</g>",
  );
}

/** 顶端对齐（单元格垂直对齐，16×16） */
export function iconAlignTop(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="currentColor" fill-rule="evenodd">' +
      '<path d="M1 0h14v1H1zM1 8h14v1H1zM3 4h10v1H3z"/>' +
      "</g>",
  );
}

/** 垂直居中（单元格垂直对齐，16×16） */
export function iconAlignMiddle(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="currentColor" fill-rule="evenodd">' +
      '<path d="M1 4h14v1H1zM1 12h14v1H1zM3 8h10v1H3z"/>' +
      "</g>",
  );
}

/** 底端对齐（单元格垂直对齐，16×16） */
export function iconAlignBottom(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="currentColor" fill-rule="evenodd">' +
      '<path d="M1 7h14v1H1zM1 15h14v1H1zM3 11h10v1H3z"/>' +
      "</g>",
  );
}

/** 自动换行（16×16，正文与箭头分层单色） */
export function iconWrapText(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" d="M6.24 7.592q-.562.498-1.082.703A3 3 0 0 1 4.04 8.5q-.984 0-1.512-.5-.53-.501-.529-1.28a1.76 1.76 0 0 1 .723-1.44q.323-.23.728-.346.298-.083.9-.158 1.226-.153 1.805-.363l.006-.276q0-.644-.287-.907-.388-.357-1.153-.357-.714 0-1.055.26t-.503.922l-.99-.14q.135-.662.445-1.069.309-.407.894-.626T4.868 2q.765 0 1.243.187t.703.472.315.717q.05.27.05.972v1.406q0 1.47.065 1.859.065.39.256.746H6.443a2.3 2.3 0 0 1-.203-.767m-.084-2.354q-.551.235-1.653.398-.624.094-.883.211a.9.9 0 0 0-.4.343.93.93 0 0 0-.14.5q0 .423.306.703.308.282.897.281.585 0 1.04-.266t.67-.73q.163-.356.163-1.053zM9.497 8.363H8.5V0h1.075v2.984a2.16 2.16 0 0 1 1.738-.816q.586 0 1.107.225.523.225.86.633.338.408.529.984.19.577.191 1.233 0 1.557-.806 2.407-.807.85-1.935.85-1.123 0-1.762-.896zm-.012-3.075q0 1.09.31 1.575.51.792 1.374.793.705 0 1.219-.585.513-.584.513-1.743 0-1.186-.492-1.751-.493-.565-1.192-.565-.705 0-1.218.585-.514.584-.514 1.691M4.771 10c.737 0 1.32.165 1.76.495.451.353.737.881.847 1.597H6.135c-.077-.364-.23-.628-.45-.804-.232-.176-.54-.264-.914-.264-.461 0-.824.165-1.088.517-.264.342-.396.826-.396 1.442 0 .628.12 1.123.385 1.476.242.341.605.517 1.088.517.836 0 1.298-.418 1.397-1.244H7.4c-.132.77-.429 1.343-.88 1.718-.44.363-1.034.55-1.77.55q-1.32 0-2.046-.859C2.23 14.591 2 13.875 2 12.994c0-.858.231-1.563.693-2.102Q3.436 10 4.77 10"/>' +
      '<path fill="currentColor" fill-opacity="0.72" fill-rule="nonzero" d="M12 13v-3h1v4h-3v1.5l-2-2 2-2V13z"/>' +
      "</g>",
  );
}

/** 合并后居中（16×16，框线与箭头分层单色） */
export function iconMerge(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-rule="nonzero" d="M8 12v3H7v-3H1v-1h13v1zM7 4V1h1v3h6v1H1V4zM1 1v14h13V1zM0 0h15v16H0z"/>' +
      '<path fill="currentColor" fill-opacity="0.72" fill-rule="nonzero" d="M8 7.5h2.5V6L13 7.997 10.5 10V8.5h-6V10L2 8.003 4.5 6v1.5z"/>' +
      "</g>",
  );
}

/** 文字方向（16×16，ab + 铅笔，单色层次） */
export function iconTextOrientation(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-opacity="0.72" fill-rule="nonzero" d="m15 5-3.793.345 1.38 1.38L5 14.31l.69.69 7.586-7.586 1.38 1.38z"/>' +
      '<path fill="currentColor" d="M6.638 9.221q-.044.715-.256 1.204a2.9 2.9 0 0 1-.612.892q-.662.663-1.353.682-.692.02-1.215-.506a1.68 1.68 0 0 1-.483-1.46q.064-.371.257-.723.145-.257.499-.714.72-.93.969-1.463l-.181-.189q-.433-.435-.803-.419-.5.021-1.015.537-.48.482-.533.887-.054.404.281.962l-.76.572q-.353-.537-.42-1.02-.064-.483.18-1.026.246-.543.764-1.062.514-.516.961-.712t.79-.156.693.271q.216.148.688.622l.944.948q.988.99 1.293 1.21T8 8.89l-.71.713a2.2 2.2 0 0 1-.652-.38m-1.639-1.53q-.186.462-.693 1.175l-.15.208q-.357.485-.452.738a.85.85 0 0 0-.038.5.9.9 0 0 0 .242.433q.283.285.678.267t.792-.415a1.93 1.93 0 0 0 .52-.882 1.55 1.55 0 0 0-.04-.943q-.13-.351-.6-.82zM9.066 6.86l-.647.64L3 2.134l.697-.69L5.63 3.36q-.088-.96.597-1.639.38-.375.864-.566.485-.19.968-.145t.98.292.922.668q1.01.999 1.038 2.062t-.703 1.787q-.728.72-1.722.555zm-2-1.965q.705.699 1.22.81.844.184 1.405-.372.456-.452.41-1.156-.045-.705-.796-1.448-.77-.762-1.454-.808-.686-.046-1.138.402-.458.453-.41 1.157.045.705.762 1.415"/>' +
      "</g>",
  );
}

/** 减少缩进量（16×16） */
export function iconDecreaseIndent(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" d="M1 0h14v1H1zM9 5h6v1H9zM9 10h6v1H9zM1 15h14v1H1z"/>' +
      '<path fill="currentColor" fill-opacity="0.72" fill-rule="nonzero" d="m1 8 3 3V8.5h4v-1H4V5z"/>' +
      "</g>",
  );
}

/** 增加缩进量（16×16） */
export function iconIncreaseIndent(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" d="M1 0h14v1H1zM9 5h6v1H9zM9 10h6v1H9zM1 15h14v1H1z"/>' +
      '<path fill="currentColor" fill-opacity="0.72" fill-rule="nonzero" d="M8 8 5 5v2.5H1v1h4V11z"/>' +
      "</g>",
  );
}

/** 逆时针旋转文字角度（16×16，与 {@link iconTextOrientation} 同形） */
export function iconRotateTextCounterClockwise(): SVGSVGElement {
  return iconTextOrientation();
}

/** 顺时针旋转文字角度（16×16） */
export function iconRotateTextClockwise(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-opacity="0.72" fill-rule="nonzero" d="m11 15-.345-3.793-1.38 1.38L1.69 5 1 5.69l7.586 7.586-1.38 1.38z"/>' +
      '<path fill="currentColor" d="M6.779 6.638q-.715-.044-1.204-.256a2.9 2.9 0 0 1-.892-.612q-.663-.662-.682-1.353-.02-.692.506-1.215a1.68 1.68 0 0 1 1.46-.483q.371.064.723.257.257.145.714.499.93.72 1.463.969l.189-.181q.435-.433.419-.803-.021-.5-.537-1.015-.482-.48-.887-.533-.404-.054-.962.281l-.572-.76q.537-.353 1.02-.42.483-.064 1.026.18.543.246 1.062.764.516.514.712.961t.156.79-.271.693q-.148.216-.622.688l-.948.944q-.99.989-1.21 1.293A2.3 2.3 0 0 0 7.11 8l-.713-.71a2.2 2.2 0 0 1 .38-.652m1.53-1.639q-.462-.186-1.175-.693l-.208-.15q-.485-.357-.738-.452a.85.85 0 0 0-.5-.038.9.9 0 0 0-.433.242.87.87 0 0 0-.267.678q.017.394.415.792.395.393.882.52.487.126.943-.04.352-.13.82-.6zM9.14 9.066l-.64-.647L13.866 3l.69.697L12.64 5.63q.96-.088 1.639.597.375.38.566.864.19.485.145.968t-.292.98a3.6 3.6 0 0 1-.668.922q-.999 1.01-2.062 1.038t-1.787-.703q-.72-.728-.555-1.722zm1.965-2q-.699.705-.81 1.22-.184.844.372 1.405.452.456 1.156.41.705-.046 1.448-.796.762-.77.808-1.454.046-.686-.402-1.138-.452-.458-1.157-.41-.705.045-1.415.762"/>' +
      "</g>",
  );
}

/** 竖排文字（16×16） */
export function iconTextVertical(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-opacity="0.72" fill-rule="nonzero" d="M4.5 15 7 12H5V1H4v11H2z"/>' +
      '<path fill="currentColor" d="M12.74 5.592q-.562.498-1.082.703a3 3 0 0 1-1.117.205q-.984 0-1.512-.5-.53-.501-.529-1.28a1.76 1.76 0 0 1 .723-1.44q.323-.23.728-.346.299-.083.9-.158 1.226-.153 1.805-.363l.006-.276q0-.644-.287-.907-.388-.357-1.153-.357-.714 0-1.055.26t-.503.922l-.99-.14q.135-.662.445-1.069.309-.407.894-.626T11.368 0q.765 0 1.243.187t.703.472.315.717q.05.27.05.972v1.406q0 1.47.065 1.859.065.39.256.746h-1.057a2.3 2.3 0 0 1-.203-.767m-.084-2.354q-.551.235-1.653.398-.625.094-.883.211a.9.9 0 0 0-.4.343.93.93 0 0 0-.14.5q0 .423.306.703.307.282.897.281.585 0 1.04-.266t.67-.73q.163-.356.163-1.053zM9.497 15.863H8.5V7.5h1.075v2.984a2.16 2.16 0 0 1 1.738-.816q.586 0 1.107.225.523.225.86.633.338.408.529.984.19.577.191 1.233 0 1.557-.806 2.407-.807.85-1.935.85-1.123 0-1.762-.896zm-.012-3.075q0 1.09.31 1.575.51.793 1.374.793.705 0 1.219-.585.513-.584.513-1.743 0-1.186-.492-1.751-.493-.565-1.192-.565-.705 0-1.218.585-.514.585-.514 1.691"/>' +
      "</g>",
  );
}

/** 向上旋转文字（16×16） */
export function iconRotateTextUp(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-opacity="0.72" fill-rule="nonzero" d="M3.5 1 1 4h2v11h1V4h2z"/>' +
      '<path fill="currentColor" d="M14.092 10.76q.498.562.703 1.082.205.521.205 1.117 0 .984-.5 1.512-.501.53-1.28.529a1.76 1.76 0 0 1-1.44-.723 2.4 2.4 0 0 1-.346-.728 7 7 0 0 1-.158-.9q-.153-1.226-.363-1.805l-.276-.006q-.644 0-.907.287-.357.388-.357 1.153 0 .714.26 1.055t.922.503l-.14.99q-.662-.135-1.069-.445-.407-.309-.626-.894t-.22-1.355q0-.765.187-1.243t.472-.703.717-.315q.27-.05.972-.05h1.406q1.47 0 1.859-.065a2.4 2.4 0 0 0 .746-.256v1.057a2.3 2.3 0 0 1-.767.203m-2.354.084q.235.551.398 1.653.094.625.211.883a.9.9 0 0 0 .343.4q.225.14.5.14a.91.91 0 0 0 .703-.306q.282-.307.281-.897 0-.585-.266-1.04a1.6 1.6 0 0 0-.73-.67q-.356-.163-1.053-.163zM14.863 5.503V6.5H6.5V5.425h2.984a2.16 2.16 0 0 1-.816-1.738q0-.586.225-1.107.225-.523.633-.86.408-.338.984-.529A3.9 3.9 0 0 1 11.743 1q1.557 0 2.407.806.85.807.85 1.935 0 1.123-.896 1.762zm-3.075.012q1.09 0 1.575-.31.792-.51.793-1.374 0-.705-.585-1.219-.584-.513-1.743-.513-1.186 0-1.751.492-.565.493-.565 1.192 0 .705.585 1.218.585.514 1.691.514"/>' +
      "</g>",
  );
}

/** 向下旋转文字（16×16） */
export function iconRotateTextDown(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-opacity="0.72" fill-rule="nonzero" d="M3.5 15 6 12H4V1H3v11H1z"/>' +
      '<path fill="currentColor" d="M7.408 5.24q-.498-.562-.703-1.082A3 3 0 0 1 6.5 3.04q0-.984.5-1.512Q7.502.999 8.28 1a1.76 1.76 0 0 1 1.44.723q.23.323.346.728.083.298.158.9.153 1.226.363 1.805l.276.006q.644 0 .907-.287.357-.388.357-1.153 0-.714-.26-1.055t-.922-.503l.14-.99q.662.135 1.069.445.407.309.626.894T13 3.868q0 .765-.187 1.243t-.472.703a1.7 1.7 0 0 1-.717.315q-.27.05-.972.05H9.246q-1.47 0-1.859.065a2.4 2.4 0 0 0-.746.256V5.443q.327-.158.767-.203m2.354-.084q-.235-.551-.398-1.653-.094-.624-.211-.883a.9.9 0 0 0-.343-.4.93.93 0 0 0-.5-.14.91.91 0 0 0-.703.306q-.282.308-.281.897 0 .585.266 1.04t.73.67q.356.163 1.053.163zM6.637 10.497V9.5H15v1.075h-2.984q.816.68.816 1.738 0 .586-.225 1.107-.225.523-.633.86-.408.338-.984.529A3.9 3.9 0 0 1 9.757 15Q8.2 15 7.35 14.194q-.85-.807-.85-1.935 0-1.123.896-1.762zm3.075-.012q-1.09 0-1.575.31-.793.51-.793 1.374 0 .705.585 1.219.584.513 1.743.513 1.186 0 1.751-.492.565-.493.565-1.192 0-.705-.585-1.218-.585-.514-1.691-.514"/>' +
      "</g>",
  );
}

/** 跨越合并（16×16） */
export function iconMergeAcross(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="nonzero">' +
      '<path fill="currentColor" d="M15 0v16H0V0zM1 12v3l6-.001V12zm13-7H1v6h13zm0-4H1v3h6V1h1v3h6zM8 14.999 14 15v-3H8z"/>' +
      '<path fill="currentColor" fill-opacity="0.72" d="M13 7.997 10.5 6v1.5H2v1h8.5V10z"/>' +
      "</g>",
  );
}

/** 合并单元格（16×16，不含居中箭头） */
export function iconMergeCells(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-rule="nonzero" d="M8 12v3H7v-3H1v-1h13v1zM7 4V1h1v3h6v1H1V4zM1 1v14h13V1zM0 0h15v16H0z"/>' +
      '<path fill="currentColor" fill-opacity="0.72" fill-rule="nonzero" d="M1 5v6h13V5zM0 4h15v8H0z"/>' +
      "</g>",
  );
}

/** 取消单元格合并（16×16） */
export function iconUnmergeCells(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-rule="nonzero" d="M8 12v3H7v-3H1v-1h13v1zM7 4V1h1v3h6v1H1V4zM1 1v14h13V1zM0 0h15v16H0z"/>' +
      '<path fill="currentColor" fill-opacity="0.72" fill-rule="nonzero" d="M1 5v6h13V5zM0 4h15v8H0zm7 1h1v6H7zM3 5h1v6H3zm8 0h1v6h-1z"/>' +
      "</g>",
  );
}

/** 百分比（16×16，单色 currentColor） */
export function iconPercent(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<path fill="currentColor" fill-rule="nonzero" d="M1 4.618q0-1.66.84-2.64Q2.68 1 4.103 1q1.366 0 2.131.9T7 4.408q0 1.6-.844 2.596T3.913 8Q2.59 8 1.794 7.074 1 6.148 1 4.618m1-.077q0 1.18.536 1.82Q3.072 7 4 7q.957 0 1.478-.663Q6 5.674 6 4.468q0-1.179-.512-1.824Q4.977 2 4.048 2 3.1 2 2.55 2.672 2 3.343 2 4.542m10.805-3.34L4.372 14.878H3.017l8.432-13.675zM9 11.632q0-1.654.841-2.643T12.1 8q1.36 0 2.13.901.771.9.771 2.538 0 1.584-.854 2.573-.853.988-2.229.988-1.335 0-2.126-.936T9 11.631m1-.08q0 1.155.536 1.802Q11.07 14 12 14q.957 0 1.478-.67.522-.67.522-1.871 0-1.193-.522-1.826Q12.958 9 12.048 9q-.976 0-1.512.674Q10 10.35 10 11.55"/>',
  );
}

/** 千分位（16×16，单色 currentColor） */
export function iconCommaStyle(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<path fill="currentColor" fill-rule="nonzero" d="m6.02 15-.971-1.306c2.576-1.62 3.743-3.187 3.743-5.067 0-.627-.195-.836-.924-1.045C5.778 7.06 5 6.172 5 4.448 5 2.463 6.361 1 8.257 1 10.542 1 12 2.88 12 5.806c0 2.09-.778 4.336-2.09 6.112-.973 1.254-2.14 2.194-3.89 3.082"/>',
  );
}

/** 增加小数位数（16×16，箭头与数字分层单色） */
export function iconIncreaseDecimal(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-opacity="0.72" fill-rule="nonzero" d="M1 3.5 3.5 6V4H8V3H3.5V1z"/>' +
      '<path fill="currentColor" d="M13 1c.636 0 1.134.275 1.494.841Q15 2.641 15 4t-.506 2.159C14.134 6.717 13.636 7 13 7q-.966-.002-1.494-.841Q11 5.359 11 4t.506-2.159C11.858 1.275 12.356 1 13 1m0 .8c-.437 0-.743.251-.92.769q-.183.521-.183 1.431 0 .897.183 1.431.265.766.92.768.644-.002.92-.768.183-.534.183-1.431 0-.91-.183-1.431c-.184-.518-.49-.768-.92-.768M13 9c.636 0 1.134.275 1.494.841Q15 10.641 15 12t-.506 2.159c-.36.558-.858.841-1.494.841q-.966-.002-1.494-.841Q11 13.359 11 12t.506-2.159C11.858 9.275 12.356 9 13 9m0 .8c-.437 0-.743.251-.92.769q-.183.521-.183 1.431 0 .897.183 1.431.265.766.92.768.644-.001.92-.768.183-.534.183-1.431 0-.91-.183-1.431c-.184-.518-.49-.768-.92-.768M7 9c.636 0 1.134.275 1.494.841Q9 10.641 9 12t-.506 2.159C8.134 14.717 7.636 15 7 15q-.966-.002-1.494-.841Q5 13.359 5 12t.506-2.159C5.858 9.275 6.356 9 7 9m0 .8c-.437 0-.743.251-.92.769q-.183.521-.183 1.431 0 .897.183 1.431.265.766.92.768.644-.001.92-.768.183-.534.183-1.431 0-.91-.183-1.431C7.736 10.05 7.43 9.8 7 9.8M8.994 5c.28 0 .51.09.713.293a.98.98 0 0 1 .293.713c0 .28-.102.51-.293.714-.204.178-.433.28-.713.28a.97.97 0 0 1-.701-.28A.95.95 0 0 1 8 6.006c0-.28.09-.522.293-.713.178-.204.42-.293.7-.293M1.994 13c.28 0 .51.09.713.293a.98.98 0 0 1 .293.713c0 .28-.102.51-.293.714-.204.178-.433.28-.713.28a.97.97 0 0 1-.701-.28.95.95 0 0 1-.293-.714c0-.28.09-.522.293-.713.178-.204.42-.293.7-.293"/>' +
      "</g>",
  );
}

/** 减少小数位数（16×16，箭头与数字分层单色） */
export function iconDecreaseDecimal(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-opacity="0.72" fill-rule="nonzero" d="M8 11.5 5.5 9v2H1v1h4.5v2z"/>' +
      '<path fill="currentColor" d="M13 1c.636 0 1.134.275 1.494.841Q15 2.641 15 4t-.506 2.159C14.134 6.717 13.636 7 13 7q-.966-.002-1.494-.841Q11 5.359 11 4t.506-2.159C11.858 1.275 12.356 1 13 1m0 .8c-.437 0-.743.251-.92.769q-.183.521-.183 1.431 0 .897.183 1.431.265.766.92.768.644-.002.92-.768.183-.534.183-1.431 0-.91-.183-1.431c-.184-.518-.49-.768-.92-.768M13 9c.636 0 1.134.275 1.494.841Q15 10.641 15 12t-.506 2.159c-.36.558-.858.841-1.494.841q-.966-.002-1.494-.841Q11 13.359 11 12t.506-2.159C11.858 9.275 12.356 9 13 9m0 .8c-.437 0-.743.251-.92.769q-.183.521-.183 1.431 0 .897.183 1.431.265.766.92.768.644-.001.92-.768.183-.534.183-1.431 0-.91-.183-1.431c-.184-.518-.49-.768-.92-.768M7 1c.636 0 1.134.275 1.494.841Q9 2.641 9 4t-.506 2.159C8.134 6.717 7.636 7 7 7q-.966-.002-1.494-.841Q5 5.359 5 4t.506-2.159C5.858 1.275 6.356 1 7 1m0 .8c-.437 0-.743.251-.92.769Q5.897 3.09 5.897 4q0 .897.183 1.431.265.766.92.768.644-.002.92-.768.183-.534.183-1.431 0-.91-.183-1.431C7.736 2.05 7.43 1.8 7 1.8M8.994 13c.28 0 .51.09.713.293a.98.98 0 0 1 .293.713c0 .28-.102.51-.293.714-.204.178-.433.28-.713.28a.97.97 0 0 1-.701-.28.95.95 0 0 1-.293-.714c0-.28.09-.522.293-.713.178-.204.42-.293.7-.293M1.994 5c.28 0 .51.09.713.293A.98.98 0 0 1 3 6.006c0 .28-.102.51-.293.714-.204.178-.433.28-.713.28a.97.97 0 0 1-.701-.28A.95.95 0 0 1 1 6.006c0-.28.09-.522.293-.713.178-.204.42-.293.7-.293"/>' +
      "</g>",
  );
}

/** 常规（16×16，单色层次） */
export function iconFormatGeneral(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-opacity="0.22" d="M7 0a7 7 0 0 1 6.71 9.001L3 9v3.745A7 7 0 0 1 7 0"/>' +
      '<path fill="currentColor" d="M5.505 10.113h.69v5.774h-.917v-4.634c-.337.316-.76.55-1.278.704v-.938q.375-.097.8-.34.423-.265.705-.566M9.338 10q.81 0 1.34.485c.345.324.517.736.517 1.254 0 .5-.188.954-.548 1.366-.22.243-.612.55-1.16.93-.573.388-.917.728-1.043 1.019h2.76v.833H7.3c0-.59.18-1.1.556-1.537.204-.242.635-.598 1.286-1.059.36-.259.611-.477.768-.647q.366-.426.368-.914c0-.315-.086-.55-.243-.703-.164-.154-.407-.226-.729-.226q-.517 0-.776.363c-.172.227-.266.575-.282 1.027h-.917c.008-.647.188-1.164.549-1.56Q8.432 10 9.338 10m4.671 0c.564 0 1.027.137 1.372.42.337.284.51.672.51 1.173 0 .63-.314 1.051-.933 1.261.329.106.587.26.76.47q.282.329.282.84 0 .8-.54 1.31c-.377.348-.87.526-1.482.526q-.872 0-1.403-.46c-.392-.34-.612-.842-.659-1.489h.933c.016.372.126.663.345.865q.294.281.776.283c.353 0 .635-.105.839-.307a.95.95 0 0 0 .274-.687c0-.332-.102-.574-.29-.728-.188-.162-.462-.234-.823-.234h-.392v-.712h.392q.495 0 .752-.218.25-.22.251-.647c0-.283-.078-.493-.227-.639q-.247-.219-.729-.218c-.33 0-.58.08-.76.25q-.282.255-.33.777h-.9c.046-.583.25-1.035.626-1.359.353-.323.807-.477 1.356-.477"/>' +
      '<path fill="currentColor" fill-rule="nonzero" d="M7 0a7 7 0 0 1 6.71 9.001h-1.052A6 6 0 1 0 3 11.473v1.272A7 7 0 0 1 7 0"/>' +
      '<path fill="currentColor" fill-rule="nonzero" d="M7 3v4h3v1H6V3z"/>' +
      "</g>",
  );
}

/** 数字（16×16，单色） */
export function iconFormatNumber(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<path fill="currentColor" fill-rule="nonzero" d="M5.014 1v14H4V3.049c-.686.63-1.65 1.426-2.351 1.827A13 13 0 0 1 0 5.686V4.517c1.166-.575 1.925-1.077 2.587-1.652C3.323 2.226 3.9 1.603 4.226 1zm7.202 0c1.16 0 2.081.361 2.759 1.094q1.008 1.092 1.008 2.706 0 .825-.297 1.62-.3.806-.997 1.698-.71.909-2.358 2.493-.784.75-1.253 1.228l-.192.199a8 8 0 0 0-.345.38c-.28.336-.67 1.038-.848 1.377L9.59 14H16V15H8c.018-.437.072-.723.166-1 .206-.627.83-1.243 1.282-1.85q.69-.925 1.996-2.14 2.061-1.924 2.785-3.049.748-1.164.746-2.2 0-1.11-.697-1.871c-.461-.504-1.355-.76-2.097-.76-.783 0-1.537.202-2.006.736-.402.458-.638 1.065-.7 1.826l-.013.212-1.158-.08c.165-1.327.54-2.195 1.157-2.808C10.145 1.336 11.064 1 12.216 1"/>',
  );
}

/** 货币（16×16，单色层次） */
export function iconFormatCurrency(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-opacity="0.22" d="M15 1v5.615C14.05 6.231 12.83 6 11.5 6 8.537 6 6.12 7.146 6.004 8.58L6 8.69V11H1V1z"/>' +
      '<path fill="currentColor" fill-rule="nonzero" d="m16 0 .001 7.372c-.315-.12-.65-.41-1-.505L15 1H1v10h5.001L6 12H0V0z"/>' +
      '<path fill="currentColor" fill-rule="nonzero" d="M5 2v1H4a1 1 0 0 0-.993.883L3 4v4a1 1 0 0 0 .883.993L4 9h1v1H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2zm3 0c1.657 0 3 1.79 3 4v.01q-.517.024-1.003.092Q10.001 6.05 10 6c0-1.657-.895-3-2-3S6 4.343 6 6c0 .642.135 1.237.364 1.725-.21.267-.335.555-.36.856L6 8.689l.001.294C5.387 8.25 5 7.185 5 6c0-2.21 1.343-4 3-4m4 0a2 2 0 0 1 2 2l.001 2.293a10 10 0 0 0-1-.192L13 4l-.007-.117A1 1 0 0 0 12 3h-1V2z"/>' +
      '<g fill-rule="nonzero">' +
      '<path fill="currentColor" fill-opacity="0.22" d="m7 9.2.004-.088C7.099 7.938 9.075 7 11.5 7c2.485 0 4.5.985 4.5 2.2v4.6c0 .974-2.015 2.2-4.5 2.2S7 14.774 7 13.8z"/>' +
      '<path fill="currentColor" fill-opacity="0.55" d="M11.5 7c2.425 0 4.401.913 4.496 2.057l.004.086v4.596C16 14.922 13.985 16 11.5 16c-2.425 0-4.401-1.031-4.496-2.175L7 13.74V9.143l.004-.086C7.099 7.913 9.075 7 11.5 7m0 7c-1.637 0-2.886-.416-3.673-1.039v.759-.138c0 .221.3.56.928.859.754.358 1.618.57 2.745.57s1.983-.212 2.737-.57c.628-.3.928-.5.928-.721v-.138l.001-.621C14.38 13.583 13.138 14 11.5 14m0-2c-1.637 0-2.885-.416-3.673-1.038v.6c0 .221.3.56.928.859.754.358 1.618.57 2.745.57s1.983-.212 2.737-.57c.628-.3.928-.478.928-.7v-.16l.001-.6C14.38 11.583 13.138 12 11.5 12m0-4.1c-2.057 0-3.704.874-3.846 1.468l-.008.053-.003.053c0 .222.3.664.928.964.754.358 1.802.57 2.929.57s2.175-.212 2.929-.57c.624-.298.924-.738.928-.96v-.024C15.303 8.819 13.618 7.9 11.5 7.9"/>' +
      '<path fill="currentColor" fill-opacity="0.38" d="M11.5 11.007c2.13 0 3.857-.764 3.857-1.553 0-.79-1.727-1.554-3.857-1.554s-3.857.765-3.857 1.554 1.727 1.553 3.857 1.553"/>' +
      "</g>" +
      "</g>",
  );
}

/** 会计专用（16×16，单色层次） */
export function iconFormatAccounting(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-opacity="0.22" d="M12 1v5.011A11 11 0 0 0 11.5 6C8.537 6 6.12 7.146 6.004 8.58L6 8.69v5.622c0 .223.086.457.247.69L2 15V1z"/>' +
      '<path fill="currentColor" fill-rule="nonzero" d="M13 0v6.101a11 11 0 0 0-1-.09V1H2v14h4.247c.24.35.648.697 1.18 1.001L1 16V0zM5 12v2H3v-2zm0-3v2H3V9zm0-3v2H3V6zm3 0v.615c-.873.352-1.52.834-1.818 1.385H6V6zm3 0v.01c-.715.032-1.39.13-2 .283V6z"/>' +
      '<path fill="currentColor" fill-rule="nonzero" d="M11 2v3H3V2zm-1 1H4v1h6z"/>' +
      '<path fill="currentColor" fill-opacity="0.45" fill-rule="nonzero" d="M4 3h6v1H4z"/>' +
      '<g fill-rule="nonzero">' +
      '<path fill="currentColor" fill-opacity="0.22" d="m7 9.2.004-.088C7.099 7.938 9.075 7 11.5 7c2.485 0 4.5.985 4.5 2.2v4.6c0 .974-2.015 2.2-4.5 2.2S7 14.774 7 13.8z"/>' +
      '<path fill="currentColor" fill-opacity="0.55" d="M11.5 7c2.425 0 4.401.913 4.496 2.057l.004.086v4.596C16 14.922 13.985 16 11.5 16c-2.425 0-4.401-1.031-4.496-2.175L7 13.74V9.143l.004-.086C7.099 7.913 9.075 7 11.5 7m0 7c-1.637 0-2.886-.416-3.673-1.039v.759-.138c0 .221.3.56.928.859.754.358 1.618.57 2.745.57s1.983-.212 2.737-.57c.628-.3.928-.5.928-.721v-.138l.001-.621C14.38 13.583 13.138 14 11.5 14m0-2c-1.637 0-2.885-.416-3.673-1.038v.6c0 .221.3.56.928.859.754.358 1.618.57 2.745.57s1.983-.212 2.737-.57c.628-.3.928-.478.928-.7v-.16l.001-.6C14.38 11.583 13.138 12 11.5 12m0-4.1c-2.057 0-3.704.874-3.846 1.468l-.008.053-.003.053c0 .222.3.664.928.964.754.358 1.802.57 2.929.57s2.175-.212 2.929-.57c.624-.298.924-.738.928-.96v-.024C15.303 8.819 13.618 7.9 11.5 7.9"/>' +
      '<path fill="currentColor" fill-opacity="0.38" d="M11.5 11.007c2.13 0 3.857-.764 3.857-1.553 0-.79-1.727-1.554-3.857-1.554s-3.857.765-3.857 1.554 1.727 1.553 3.857 1.553"/>' +
      "</g>" +
      "</g>",
  );
}

/** 日期（16×16，单色层次） */
export function iconFormatDate(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="nonzero">' +
      '<path fill="currentColor" fill-opacity="0.22" d="M1 2h14v13H1z"/>' +
      '<path fill="currentColor" d="M11 2h4v13H1V2h4V1h1v1h4V1h1zm0 1v1h-1V3H6v1H5V3H2v11h12V3zM2 5h12v1H2z"/>' +
      '<path fill="currentColor" fill-opacity="0.72" d="M4 8h2v2H4z"/>' +
      "</g>",
  );
}

/** 时间（16×16，单色层次） */
export function iconFormatTime(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="none" fill-rule="evenodd">' +
      '<circle cx="8" cy="8" r="8" fill="currentColor" fill-opacity="0.22"/>' +
      '<path fill="currentColor" fill-rule="nonzero" d="M8 0a8 8 0 1 1 0 16A8 8 0 0 1 8 0m0 1a7 7 0 1 0 0 14A7 7 0 0 0 8 1"/>' +
      '<path fill="currentColor" fill-rule="nonzero" d="M8 3v5h3v1H7V3z"/>' +
      "</g>",
  );
}

/** 分数（16×16，单色） */
export function iconFormatFraction(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="currentColor" fill-rule="evenodd">' +
      '<path fill-rule="nonzero" d="M12.805 1.202 4.372 14.877H3.017l8.432-13.675z"/>' +
      '<path d="M3.198 1.113h.836v7H2.923V2.496c-.41.382-.922.666-1.55.853V2.21q.457-.117.97-.412c.342-.215.627-.44.855-.686M12.723 8q.964 0 1.594.577c.41.385.615.875.615 1.49 0 .596-.224 1.135-.652 1.625-.261.289-.727.654-1.38 1.106-.68.462-1.09.865-1.24 1.212h3.281V15H10.3c0-.702.214-1.308.662-1.827.242-.288.754-.711 1.528-1.26.429-.307.727-.567.913-.769.29-.336.438-.702.438-1.086 0-.375-.102-.654-.288-.837-.196-.183-.485-.27-.867-.27q-.615 0-.923.434c-.205.269-.317.682-.335 1.22h-1.09q.01-1.15.651-1.855.658-.75 1.734-.75"/>' +
      "</g>",
  );
}

/** 科学计数（16×16，单色） */
export function iconFormatScientific(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<g fill="currentColor" fill-rule="evenodd">' +
      '<path fill-rule="nonzero" d="M4.014 4v11H3V6.049C2.314 6.679.767 7.646.27 7.84V6.674c.643-.33 1.197-.635 1.663-1.1.73-.728.967-.97 1.293-1.574z"/>' +
      '<path d="M14.03 1q.69 0 1.14.412c.292.275.438.625.438 1.065 0 .425-.16.81-.466 1.16-.186.206-.519.467-.985.79q-.728.495-.885.866h2.343V6H12.3c0-.501.153-.934.472-1.305q.262-.31 1.092-.9c.307-.22.52-.405.653-.55.206-.24.313-.5.313-.775 0-.268-.074-.467-.207-.598q-.21-.194-.619-.192-.44 0-.659.309c-.146.192-.226.488-.24.872h-.779c.007-.55.16-.989.466-1.325Q13.262 1 14.031 1"/>' +
      '<path fill-rule="nonzero" d="M8.5 4q.86 0 1.508.338.644.34 1.065.977.432.657.678 1.601.25.96.249 2.588.001 1.925-.4 3.107-.389 1.147-1.164 1.775-.766.616-1.936.614c-1.026 0-1.835-.358-2.421-1.083C5.351 13.014 5 11.54 5 9.504q-.001-1.94.404-3.122.392-1.147 1.167-1.771C7.082 4.2 7.726 4 8.5 4m-.016 1q-1.075 0-1.718.833C6.265 6.496 6 7.713 6 9.5c0 1.79.244 2.972.7 3.563.486.63 1.09.937 1.8.937s1.314-.308 1.8-.941c.456-.593.7-1.775.7-3.559 0-1.794-.244-2.978-.7-3.566C9.814 5.306 9.206 5 8.484 5"/>' +
      "</g>",
  );
}

/** 文本（16×16，单色） */
export function iconFormatText(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<path fill="currentColor" fill-rule="nonzero" d="M3.64 6.136q.98-.001 1.592.255c.39.163.678.364.862.612q.274.372.382.934.059.34.064 1.175v2.174c.003 1.506.034 2.326.089 2.681q.053.345.167.67l.061.163H5.853a3.5 3.5 0 0 1-.21-.917 4 4 0 0 1-.018-.401 7 7 0 0 1-.273.272q-.72.685-1.389.967Q3.3 15 2.54 15c-.819 0-1.453-.215-1.894-.662q-.647-.656-.646-1.676c0-.4.08-.968.244-1.298q.244-.496.64-.795.409-.308.915-.465.39-.115 1.174-.22 1.657-.222 2.44-.527c.126-.178.126-.522.126-.578 0-.689-.158-.957-.441-1.236-.371-.365-.724-.558-1.457-.558-.684 0-1.225.237-1.55.503-.272.223-.646.493-.802 1.003l-.046.157-.966-.147c.143-.576.3-.942.508-1.234q.385-.539 1.112-.83.752-.3 1.744-.3M9.537 1v6.488c.802-1.124 1.806-1.59 2.839-1.59q.789 0 1.494.32.7.318 1.15.893.459.586.718 1.413A6 6 0 0 1 16 10.312c0 1.497-.36 2.66-1.095 3.478Q13.82 14.999 12.298 15c-1.17 0-2.22-.58-2.798-1.738l-.074-.154v1.687h-.855V1zm2.957 5.985c-.757 0-1.415.316-1.967.958-.535.621-.812 1.26-.812 2.438 0 1.153.17 1.769.492 2.28.546.871 1.287 1.405 2.216 1.405.757 0 1.273-.316 1.826-.958.534-.62.811-1.435.811-2.667 0-1.261-.268-1.929-.78-2.528-.531-.621-1.034-.928-1.786-.928m-6.95 3.288c-.517.237-1.598.373-2.684.536-.622.095-.926.204-1.184.322q-.426.196-.657.573-.233.377-.233.837c0 .47.035.86.371 1.173.323.299.791.459 1.415.459.595 0 1.124-.21 1.588-.484.472-.28.91-.663 1.13-1.148.168-.37.256-.848.254-1.575z"/>',
  );
}

export function iconChart(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/>',
  );
}

export function iconTable(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18"/><path d="M3 15h18"/><path d="M9 3v18"/>',
  );
}

export function iconImage(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><path d="M21 15l-5-5L5 21"/>',
  );
}

export function iconShapes(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<circle cx="8" cy="8" r="3"/><rect x="13" y="5" width="6" height="6" rx="1"/><path d="M3 21l7-7 4 4 7-7"/>',
  );
}

export function iconSparkline(): SVGSVGElement {
  return svgEl("0 0 24 24", '<path d="M3 17l6-6 4 4 8-8"/><path d="M14 7h7v7"/>');
}

/** 数据透视表选项（32×32，单色层次） */
export function iconPivotTableOption(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-opacity="0.08" d="M3 3h26v26H3z"/>' +
      '<path fill="currentColor" d="M30 2v28H2V2zm-1 1H3v26h26z"/>' +
      '<path fill="currentColor" fill-opacity="0.62" d="M13 5h14v6H13zM5 5h6v6H5zM11 13v14H5V13z"/>' +
      '<path fill="currentColor" fill-opacity="0.34" d="M14 6h12v4H14zM6 6h4v4H6zM10 14v12H6V14z"/>' +
      '<path fill="currentColor" fill-opacity="0.88" d="m24.5 13 2.5 3h-2v9h-8v2l-3-2.5 3-2.5v2h7v-8h-2z"/>' +
      "</g>",
  );
}

/** 插入表格（32×32，单色；结构同表格式图标，无固定品牌色） */
export function iconInsertTable(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-opacity="0.08" d="M3 4h26v23H3z"/>' +
      '<path fill="currentColor" d="M30 3v25H2V3zm-1 1H3v23h26z"/>' +
      '<path fill="currentColor" fill-opacity="0.48" d="M10 8v6h6V8h1v6h6V8h1v6h5v1h-5v6h5v1h-5v5h-1v-5h-6v5h-1v-5h-6v5H9v-5H3v-1h6v-6H3v-1h6V8zm6 7h-6v6h6zm7 0h-6v6h6z"/>' +
      '<path fill="currentColor" fill-opacity="0.78" d="M2 3h28v6H2z"/>' +
      "</g>",
  );
}

/** 插入图片（32×32，画框 + 山景 + 圆日；层次用 fill-opacity，无固定 chroma） */
export function iconInsertPicture(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-opacity="0.08" d="M3 4h26v23H3z"/>' +
      '<path fill="currentColor" d="M30 3v25H2V3zm-1 1H3v23h26z"/>' +
      '<path fill="currentColor" fill-opacity="0.3" d="M3 21V11.4l7.09 7.38 9.3-10.26L27 18v3H3z"/>' +
      '<path fill="currentColor" fill-opacity="0.52" d="m3 21 5.12-5.72 9.22 10.04 4.87-5.22L27 17.03V21H3"/>' +
      '<circle cx="21" cy="10" r="2" fill="currentColor" fill-opacity="0.62"/>' +
      "</g>",
  );
}

/** 图片格式 — 更正（32×32，原稿黄/橙改为 currentColor 层次） */
export function iconPictureCorrect(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    '<g fill="none" fill-rule="evenodd">' +
      '<g transform="translate(2.948 2.754)">' +
      '<circle cx="13.052" cy="13.246" r="7" fill="currentColor" fill-opacity="0.2"/>' +
      '<path fill="currentColor" fill-rule="nonzero" d="M13.052 6.246a7 7 0 1 1 0 14 7 7 0 0 1 0-14m0 1a6 6 0 1 0 0 12 6 6 0 0 0 0-12M10.252 21.802l-1.916 4.664-.925-.38 1.91-4.648q.3.138.615.254zm6.879-.531 1.996 4.782-.923.386-1.99-4.765a9 9 0 0 0 .917-.403M4.976 17.223.413 19.127l-.385-.923 4.556-1.903q.171.474.392.922m16.606-1.1 4.884 2.008-.38.925-4.876-2.004q.14-.3.26-.613zm4.471-8.784.386.923-4.88 2.037a9 9 0 0 0-.38-.926zM.38 7.41l4.583 1.884a9 9 0 0 0-.388.92L0 8.336zM18.13 0l.926.38-1.979 4.814a9 9 0 0 0-.918-.397zM8.263.028l1.951 4.675a9 9 0 0 0-.93.368L7.338.413z"/>' +
      "</g>" +
      "</g>",
  );
}

/** 图片格式 — 颜色（32×32，画框 + 图像；原蓝/白改为单色层次） */
export function iconPictureColor(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    '<g fill="none" fill-rule="nonzero">' +
      '<path fill="currentColor" d="M30 4v24H2V4zm-1 1H3v22h26z"/>' +
      '<path fill="currentColor" fill-opacity="0.1" d="M3 5h26v22H3z"/>' +
      '<path fill="currentColor" fill-opacity="0.34" d="M3 27v-7.638l7.092-7.375 9.297 10.26V22l4.58-5L29 21.378V27h-9.61z"/>' +
      '<path fill="currentColor" fill-opacity="0.72" d="m10.12 11.283 9.218 10.039 4.874-5.216L29 21.026v1.315l-4.822-4.733-4.161 4.452L24.554 27h-1.271L10.137 12.717 3 20.2v-1.449z"/>' +
      '<circle cx="23" cy="10" r="2" fill="currentColor" fill-opacity="0.5"/>' +
      '<path fill="currentColor" fill-opacity="0.14" d="M2 4h28v24H2z"/>' +
      "</g>",
  );
}

/** 图片格式 — 透明度（32×32，左右对比 + 竖分割线；原多色改为单色层次） */
export function iconPictureTransparency(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" fill-rule="nonzero" d="M14 4v1H3v22h11v1H2V4zm16 0v24H17v-1h12V5H17V4z"/>' +
      '<path fill="currentColor" fill-opacity="0.12" fill-rule="nonzero" d="M14 5v22H3V5zm15 0v22H17V5z"/>' +
      '<path fill="currentColor" fill-opacity="0.38" fill-rule="nonzero" d="m16.983 19.592 2.406 2.655V22l4.58-5L29 21.378V27H17.025z"/>' +
      '<path fill="currentColor" fill-opacity="0.28" fill-rule="nonzero" d="M3 27v-7.638l7.092-7.375 3.927 4.334L13.997 27z"/>' +
      '<path fill="currentColor" fill-opacity="0.62" fill-rule="nonzero" d="M10.12 11.283 14 15.508v1.406l-3.863-4.197L3 20.2v-1.449z"/>' +
      '<circle cx="23" cy="10" r="2" fill="currentColor" fill-opacity="0.45" fill-rule="nonzero"/>' +
      '<path fill="currentColor" d="M15 2h1v28h-1z"/>' +
      '<path fill="currentColor" fill-opacity="0.55" fill-rule="nonzero" d="m24 16.59 5 4.788v1.315l-5.033-4.601-4.037 4.661L23.725 27h-1.222L17 20.783v-1.397l2.277 2.63z"/>' +
      "</g>",
  );
}

/** 图片格式 — 更改图片（32×32，画框 + 裁切角标；单色层次） */
export function iconPictureChange(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    '<g fill="none" fill-rule="nonzero">' +
      '<path fill="currentColor" d="M1.999 25.742c.307.254.643.475 1 .658L3 27h3.859l-.99 1H2ZM30 4v24H10.092l.99-1H29V5H3v11.6a5.5 5.5 0 0 0-1.001.658L2 4z"/>' +
      '<path fill="currentColor" fill-opacity="0.12" d="M3 26.4c.61.312 1.285.513 2 .578V27h1.86H3ZM29 5v22H11.083l1.558-1.573-5.555-5.608-2.122 2.121L7.004 24H5v-.05a2.5 2.5 0 0 1 0-4.9v-3.028a5.5 5.5 0 0 0-2 .578V5z"/>' +
      '<path fill="currentColor" fill-opacity="0.36" d="M3 26.4c.61.312 1.285.513 2 .578V27H3Zm7.092-14.413 9.297 10.26V22l4.58-5L29 21.378V27H11.084l1.557-1.573-5.555-5.608-2.122 2.121L7.004 24H5v-.05a2.5 2.5 0 0 1 0-4.9V17.28z"/>' +
      '<path fill="currentColor" fill-opacity="0.68" d="m10.12 11.283 9.218 10.039 4.874-5.216L29 21.026v1.315l-4.822-4.733-4.161 4.452L24.554 27h-1.271L10.137 12.717 5 18.103v-1.45z"/>' +
      '<circle cx="23" cy="10" r="2" fill="currentColor" fill-opacity="0.48"/>' +
      '<path fill="currentColor" fill-opacity="0.72" d="m11 25.508-3.842 4-.658-.68 2.72-2.826c-.523-.005-2.954-.008-3.292-.008h-.382q-.553.001-.553.014c-4.316 0-5.82-7.41-.967-8.904v1.03C.504 19.721 2 25.008 5.01 25.008h4.204L6.5 22.188l.658-.68z"/>' +
      "</g>",
  );
}

/** 图片格式 — 重置图片（32×32，画框 + 复位箭头；单色层次） */
export function iconPictureReset(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    '<g fill="none" fill-rule="nonzero">' +
      '<path fill="currentColor" d="M30 4v24H2V8.234l1 1.01V27h26V5H9.67a5.47 5.47 0 0 0-2.674-.985V4z"/>' +
      '<path fill="currentColor" fill-opacity="0.12" d="M29 5v22H3V9.244l1.91 1.93 2.122-2.122-2.04-2.06h2.004v.05a2.5 2.5 0 0 1 0 4.9v3.028a5.5 5.5 0 0 0 2.673-9.971z"/>' +
      '<path fill="currentColor" fill-opacity="0.36" d="m10.862 12.838 8.527 9.41V22l4.58-5L29 21.378V27H3v-7.638l4.253-4.421a5.5 5.5 0 0 0 3.61-2.103"/>' +
      '<path fill="currentColor" fill-opacity="0.68" d="m11.158 12.413 8.18 8.909 4.874-5.216L29 21.026v1.315l-4.822-4.733-4.161 4.452L24.554 27h-1.271L10.57 13.187q.326-.36.587-.774M6.996 14.56v.41q.646-.06 1.24-.259L3 20.2v-1.449z"/>' +
      '<circle cx="23" cy="10" r="2" fill="currentColor" fill-opacity="0.48"/>' +
      '<path fill="currentColor" fill-opacity="0.72" d="m1.01 5.494 3.842-4 .658.679L2.79 5c.523.004 2.954.007 3.292.008h.382q.553-.002.553-.014c4.316 0 5.82 7.41.967 8.904v-1.03C11.505 11.28 10.01 5.993 7 5.993H2.796l2.714 2.82-.658.68z"/>' +
      "</g>",
  );
}

export function iconOrientation(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<rect x="4" y="4" width="16" height="10" rx="1"/><path d="M8 20h8"/><path d="M12 14v6"/>',
  );
}

export function iconMargins(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<rect x="6" y="6" width="12" height="12"/><path d="M3 3v18"/><path d="M21 3v18"/><path d="M3 3h18"/><path d="M3 21h18"/>',
  );
}

export function iconBackground(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 15l6-6 4 4 8-8"/>',
  );
}

export function iconPrintTitles(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<rect x="4" y="4" width="16" height="16" rx="1"/><path d="M4 8h16"/><path d="M8 4v4"/>',
  );
}

export function iconFunction(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<path fill="currentColor" fill-rule="nonzero" d="M1 1h14v4.375h-1.077L12.846 2.75H6.45l3.866 4.489-4.426 5.136h6.957l1.077-1.75H15V15H1l6.03-7z"/>',
  );
}

/** Ribbon「插入函数」：32×32 fx，单色 `currentColor`。 */
export function iconInsertFunction(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    '<path fill="currentColor" fill-rule="nonzero" d="M19.433 3c.757 0 1.36.194 1.817.567.481.39.708.817.708 1.27q-.002.567-.37.934-.369.37-.946.37c-.88-.96-1.41-1.464-1.807-1.716-1.603.508-2.143 1.24-2.62 2.227q-.371.772-1.327 3.807h2.079l-.356 1.23h-2.046l-1.527 5.899Q12 21.576 10.96 23.778c-1.009 2.074-2.092 3.502-3.24 4.285q-1.343.939-2.685.937-.924-.002-1.535-.532-.5-.39-.5-1.014-.002-.51.412-.89.403-.365.975-.366c.644.835 1.013 1.267 1.302 1.492 1.211-.2 1.586-.541 1.913-.997.343-.48.667-1.176.976-2.077.128-.389.478-1.655 1.049-3.807l2.399-9.12H9.584l.26-1.23c.827.005 1.399-.042 1.715-.153q.428-.15.782-.563c.251-.289.581-.84.993-1.643.555-1.11 1.095-1.966 1.615-2.579q1.056-1.258 2.237-1.888C17.983 3.208 18.732 3 19.433 3m1.83 7.092c.504.515.883 1.037 1.133 1.548q.245.463.761 2.3l1.05-1.532c.336-.458.74-.886 1.217-1.292.486-.418.917-.691 1.293-.834 1.287-.176 1.614-.03 1.87.201.278.25.413.547.413.887-.545 1.145-.893 1.223-1.263 1.223-1.092-.264-1.377-.34-1.526-.34-1.27.803-1.939 1.697-2.72 3.01l1.144 4.678c.139.577.252.964.354 1.153 1.414-.493 1.84-1.012 2.187-1.637.19 1.458-.586 2.4-1.532 3.145-.55.428-1.023.632-1.41.632-1.05-.394-1.234-.633-1.402-.938-.193-.35-.578-1.776-1.155-4.287-1.378 2.315-2.498 3.814-3.352 4.505-.607.484-1.197.72-1.764.72-.412 0-.79-.14-1.143-.437-.253-.233-.385-.545-.385-.94q.001-.511.354-.864.352-.352.876-.354c.317 0 .655.137 1.01.433 1.338-.254 1.903-.988 2.621-2.082q1.092-1.664 1.455-2.434-.745-2.886-.884-3.38l-.027-.09c-.208-.552-.47-.953-.8-1.182-.505-.353-1.363-.414-2.555-.248z"/>',
  );
}

/** Ribbon「公式」资源：账本底图 + 符号；层次用 `fill-opacity`，无固定 chroma。 */
const FORMULA_RIBBON_BOOK_MONO =
  '<path fill="currentColor" fill-opacity="0.22" fill-rule="nonzero" d="M27 23.437q0 .875-.737 1.203a9.6 9.6 0 0 0 0 3.61h.148q.59 0 .589.656v.438q0 .656-.59.656H8.144q-1.277 0-2.21-1.04Q5 27.924 5 26.5V6.375q0-1.805 1.154-3.09T8.929 2h16.893q.49 0 .834.383.345.383.344.93z"/>' +
  '<path fill="currentColor" fill-rule="nonzero" d="M27 23.808c0 .583-.33.984-.822 1.203-.23 1.167-.23 2.75 0 3.99h.147q.59 0 .675.44c0 .414-.196.559-.59.559H8.144c-.852 0-1.772-.305-2.395-.998C5.127 28.309 5 27.449 5 26.5V6.375q0-1.805 1.154-3.09T8.929 2h16.893q.49 0 .834.383.345.383.344.93zm-1.871 5.194c-.161-1.167-.161-2.822 0-3.989H8.119c-.419 0-1.125.133-1.537.565s-.568.871-.568 1.43c0 .557.068.922.374 1.268.307.346.821.726 1.175.726zm.864-4.977V3H8.93c-.655 0-1.552.372-2.13 1.029-.58.656-.785 1.616-.785 2.346v18.337c.322-.388 1.017-.687 1.572-.687z"/>';

/** Ribbon「公式 → 自动求和」Σ（与 Office 资源同路径，单色）。 */
export function iconFormulaRibbonAutoSum(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    '<path fill="currentColor" fill-rule="nonzero" d="M3 3h26v6.125h-1L27 5.25H11.12l8.18 9.336L9.08 26.125H27l1-2.25h1V29H3l11.2-13z"/>',
  );
}

/** 公式库 · 财务（账本 + 硬币堆叠），单色。 */
export function iconFormulaRibbonFinancial(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    FORMULA_RIBBON_BOOK_MONO +
      '<g transform="translate(10 7)">' +
      '<ellipse cx="6" cy="2.5" fill="currentColor" fill-opacity="0.45" rx="6" ry="2.5"/>' +
      '<path fill="currentColor" fill-rule="nonzero" d="M6 0C3.11 0 0 .871 0 2.786v8.143c0 1.913 3.11 2.785 6 2.785s6-.872 6-2.785V2.786C12 .87 8.89 0 6 0m5.143 10.929c0 1.064-2.303 1.928-5.143 1.928S.857 11.993.857 10.93V9.327c.886.913 3.022 1.387 5.143 1.387s4.257-.474 5.143-1.387zm0-2.572h-.002l.002.013c0 1.059-2.303 1.916-5.143 1.916S.857 9.429.857 8.37l.002-.013H.857V6.756C1.743 7.668 3.88 8.143 6 8.143s4.257-.475 5.143-1.387zm0-2.571h-.002l.002.013c0 1.058-2.303 1.915-5.143 1.915S.857 6.857.857 5.8l.002-.013H.857V4.313C1.98 5.17 4.04 5.571 6 5.571s4.02-.402 5.143-1.258zM6 4.714C3.16 4.714.857 3.85.857 2.786.857 1.72 3.16.857 6 .857s5.143.863 5.143 1.929C11.143 3.85 8.84 4.714 6 4.714"/>' +
      "</g>",
  );
}

/** 公式库 · 逻辑（账本 + 问号），单色。 */
export function iconFormulaRibbonLogical(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    FORMULA_RIBBON_BOOK_MONO +
      '<path fill="currentColor" fill-rule="nonzero" d="M16.54 6c1.28 0 2.32.34 3.08 1.06.74.68 1.12 1.62 1.12 2.82 0 .9-.26 1.68-.74 2.34-.2.24-.76.76-1.64 1.54-.42.36-.72.72-.92 1.08-.26.44-.38.94-.38 1.48v.46h-1.62v-.46c0-.66.12-1.24.38-1.74.26-.6.94-1.36 2.02-2.32.32-.32.54-.56.68-.72.38-.5.58-1.02.58-1.58q0-1.2-.66-1.86c-.48-.48-1.14-.7-1.98-.7-.98 0-1.72.32-2.22.98-.42.56-.64 1.32-.64 2.3H12c0-1.4.4-2.52 1.18-3.36Q14.41 6 16.54 6m-.28 12.18c.34 0 .62.1.86.34.22.22.34.5.34.84 0 .36-.12.64-.34.86-.24.22-.54.34-.86.34-.36 0-.64-.12-.86-.34-.24-.24-.36-.52-.36-.86q0-.51.36-.84c.22-.24.5-.34.86-.34"/>',
  );
}

/** 公式库 · 文本（账本 + A），单色。 */
export function iconFormulaRibbonText(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    FORMULA_RIBBON_BOOK_MONO +
      '<path fill="currentColor" fill-rule="nonzero" d="M15.62 6h1.9l5.62 14.28h-1.78l-1.52-4h-6.56l-1.52 4H10zm-1.82 8.9h5.52l-2.7-7.14h-.08z"/>',
  );
}

/** 公式库 · 日期和时间（账本 + 时钟），单色。 */
export function iconFormulaRibbonDateTime(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    FORMULA_RIBBON_BOOK_MONO +
      '<path fill="currentColor" fill-rule="nonzero" d="M16 7c-3.86 0-7 3.14-7 7s3.14 7 7 7 7-3.14 7-7-3.14-7-7-7m0 13c-3.308 0-6-2.692-6-6s2.692-6 6-6 6 2.692 6 6-2.692 6-6 6m.412-9.734h-.824v4.146h4.277v-.824h-3.453z"/>',
  );
}

/** 公式库 · 查找与引用（账本 + 放大镜），单色。 */
export function iconFormulaRibbonLookup(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    FORMULA_RIBBON_BOOK_MONO +
      '<path fill="currentColor" fill-rule="nonzero" d="M14.912 7a5.9 5.9 0 0 1 4.19 1.722 5.9 5.9 0 0 1 1.722 4.19 5.92 5.92 0 0 1-1.35 3.77l3.351 3.352a.536.536 0 0 1 0 .791.536.536 0 0 1-.79 0l-3.353-3.351a5.9 5.9 0 0 1-3.724 1.35 5.9 5.9 0 0 1-4.19-1.723C9.699 16.031 9 14.541 9 12.911a5.9 5.9 0 0 1 1.722-4.189A5.9 5.9 0 0 1 14.912 7m3.398 2.514a4.7 4.7 0 0 0-3.398-1.397 4.79 4.79 0 0 0-3.398 1.397 4.7 4.7 0 0 0-1.397 3.398c0 1.35.512 2.514 1.397 3.398a4.68 4.68 0 0 0 3.351 1.397c1.304 0 2.514-.512 3.352-1.397a4.68 4.68 0 0 0 1.396-3.352c.094-1.396-.465-2.56-1.303-3.444"/>',
  );
}

/** 公式库 · 数学和三角函数（账本 + 圆柱分段示意），单色。 */
export function iconFormulaRibbonMath(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    FORMULA_RIBBON_BOOK_MONO +
      '<path fill="currentColor" fill-rule="nonzero" d="M16 6c1.834 0 3.166.841 3.996 2.566C20.65 9.903 21 11.714 21 14c0 2.243-.35 4.054-1.004 5.434C19.166 21.137 17.834 22 16 22c-1.769 0-3.079-.841-3.93-2.501C11.35 18.119 11 16.286 11 14c0-2.307.35-4.14 1.07-5.52C12.92 6.82 14.23 6 16 6m0 1.445c-1.179 0-2.052.733-2.598 2.2q-.524 1.358-.59 3.492h6.398c-.066-1.423-.262-2.587-.612-3.493-.546-1.466-1.419-2.2-2.598-2.2m3.231 7.137h-6.44c.043 1.51.24 2.782.61 3.774.547 1.466 1.42 2.2 2.599 2.2s2.052-.755 2.598-2.222c.371-.992.59-2.242.633-3.752"/>',
  );
}

/** 公式库 · 其他函数（账本 + 三点），单色。 */
export function iconFormulaRibbonMore(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    FORMULA_RIBBON_BOOK_MONO +
      '<circle cx="11.5" cy="13.5" r="1.5" fill="currentColor"/><circle cx="16.5" cy="13.5" r="1.5" fill="currentColor"/><circle cx="21.5" cy="13.5" r="1.5" fill="currentColor"/>',
  );
}

/** 清除（16×16，单色） */
export function iconClear(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<path fill="currentColor" fill-rule="evenodd" d="m2.721 7.612 5.667 5.667-.935.935a2.685 2.685 0 0 1-3.797 0l-1.87-1.87a2.685 2.685 0 0 1 0-3.797zm.945-.945L9.333 1 15 6.667l-5.667 5.667z"/>',
  );
}

export function iconAudit(): SVGSVGElement {
  return svgEl("0 0 24 24", '<circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/>');
}

export function iconSort(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<path d="M3 6h7"/><path d="M3 12h5"/><path d="M3 18h3"/><path d="M16 5v14l4-4"/>',
  );
}

export function iconFilter(): SVGSVGElement {
  return svgEl("0 0 24 24", '<path d="M4 4h16l-6 8v6l-4 2v-8z"/>');
}

export function iconSlicer(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<rect x="4" y="4" width="16" height="6" rx="1"/><rect x="4" y="14" width="10" height="6" rx="1"/><circle cx="17" cy="17" r="2"/>',
  );
}

export function iconDataTools(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/>',
  );
}

export function iconNormalView(): SVGSVGElement {
  return svgEl("0 0 24 24", '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18"/>');
}

export function iconPageBreak(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<rect x="3" y="3" width="18" height="18" rx="1"/><path d="M12 3v18"/><path d="M3 12h18" stroke-dasharray="2 2"/>',
  );
}

export function iconGrid(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<rect x="3" y="3" width="18" height="18"/><path d="M3 9h18M3 15h18M9 3v18M15 3v18"/>',
  );
}

export function iconZoom(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<circle cx="11" cy="11" r="8"/><path d="M21 21l-4.35-4.35"/><path d="M11 8v6M8 11h6"/>',
  );
}

export function iconFreeze(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<rect x="3" y="3" width="18" height="18"/><path d="M9 3v18"/><path d="M3 9h18"/>',
  );
}

export function iconWindow(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<rect x="3" y="3" width="8" height="8" rx="1"/><rect x="13" y="13" width="8" height="8" rx="1"/><path d="M13 3h8v8"/>',
  );
}

export function iconRuler(): SVGSVGElement {
  return svgEl("0 0 24 24", '<path d="M4 20h16"/><path d="M5 20v-5M9 20v-3M13 20v-5M17 20v-3"/>');
}

export function iconFormulaBar(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<rect x="3" y="7" width="18" height="10" rx="1"/><path d="M7 12h10"/><path d="M5 5h4"/><path d="M15 5h4"/>',
  );
}

export function iconRecordMacro(): SVGSVGElement {
  return svgEl("0 0 24 24", '<circle cx="12" cy="12" r="5"/><circle cx="12" cy="12" r="2.5"/>');
}

export function iconRelativeRef(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<path d="M4 12h16"/><path d="M8 8l-4 4 4 4"/><path d="M16 16l4-4-4-4"/>',
  );
}

/** 撤销（16×16 实心路径） */
export function iconUndo(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<path fill="currentColor" fill-rule="nonzero" d="M4.234 8.048A.24.24 0 0 1 4.09 8L.1 4.715a.243.243 0 0 1-.052-.34l.052-.05 3.993-3.278a.243.243 0 0 1 .385.194V4.02h6.065a5.458 5.458 0 0 1 0 10.916H6.077v-1.213h4.465a4.245 4.245 0 1 0 0-8.49H4.477v2.573c0 .134-.11.242-.243.242"/>',
  );
}

/** 重做（16×16 实心路径） */
export function iconRedo(): SVGSVGElement {
  return svgElMarkup(
    "0 0 16 16",
    '<path fill="currentColor" fill-rule="nonzero" d="M11.766 8.048A.24.24 0 0 0 11.91 8L15.9 4.715a.243.243 0 0 0 .052-.34l-.052-.05-3.993-3.278a.243.243 0 0 0-.385.194V4.02H5.458a5.458 5.458 0 0 0 0 10.916h4.465v-1.213H5.458a4.245 4.245 0 1 1 0-8.49h6.065v2.573c0 .134.11.242.243.242"/>',
  );
}

/**
 * 条件格式（32×32，与 Office 资源同几何；原稿中的黄/蓝等改为 currentColor + fill-opacity，无固定色）。
 */
export function iconConditionalFormatting(): SVGSVGElement {
  const svg = svgElMarkup(
    "0 0 32 32",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" d="M22 27h-7V16h14v11zH1V2h28v25z"/>' +
      '<path fill="currentColor" fill-opacity="0.12" d="M28 16H15v10H2V3h26z"/>' +
      '<path fill="currentColor" fill-opacity="0.42" d="M15 21H9v5H8v-5H2v-1h6v-5H2v-1h6V9H2V8h6V3h1v5h12V3h1v5h6v1h-6v5h6v1h-6v1h-1v-1H9v5h6zm6-7V9H9v5z"/>' +
      '<path fill="currentColor" fill-opacity="0.38" d="M9 3h8v5H9zM21 16h-6v4H9v-5h12z"/>' +
      '<path fill="currentColor" fill-opacity="0.62" d="M9 9h11v5H9zM9 21h3v5H9z"/>' +
      '<path fill="currentColor" d="M16 17h15v13H16z"/>' +
      '<path fill="currentColor" fill-opacity="0.12" d="M17 18h13v11H17z"/>' +
      '<path fill="currentColor" d="m23.747 22 1.116-2 .887.53-.821 1.47H27v.992h-2.625l-.562 1.006H27v.996h-3.743L22.137 27l-.887-.53.825-1.476H20v-.996h2.631l.562-1.006H20V22z"/>' +
      "</g>",
  );
  /* 略浅于默认字色，与其它 Ribbon 大图标观感接近 */
  svg.setAttribute("opacity", "0.68");
  return svg;
}

/** 套用表格样式（32×32，表格 + 格式刷，单色层次） */
export function iconTableStyle(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" d="m29 2 .001 7.338q-.474.196-1 .478L28 3H2v23h6.008a3 3 0 0 0-.22.217c-.234.254-.381.519-.457.783H1V2zm0 14.768V27h-8.304q.266-.574.391-1.136l-.034.136H28v-7.87q.528-.69 1-1.362"/>' +
      '<path fill="currentColor" fill-opacity="0.12" d="M28 3v6.816c-2.48 1.327-5.735 4.029-8.678 7.184-1.037 1.112-1.966 2.096-2.772 3.004-1.256-.053-2.491.466-3.468 1.409-.52.501-.882 1.047-1.318 1.869l-.076.145-.099.189-.098.187c-.194.37-.318.578-.428.716.015-.018-.322.157-.975.416l-.4.158c-.811.326-1.285.555-1.68.907H2V3zm0 15.13V26h-6.947l.034-.136c.9-.953 2.016-2.14 3.373-3.595l.345-.369A51 51 0 0 0 28 18.13"/>' +
      '<path fill="currentColor" fill-opacity="0.42" d="M23 23.83V26h-1v-1.103zM9 3v5h6V3h1v5h6V3h1v5h5v1h-5v4.432c-.567.496-1.14 1.02-1.71 1.569L16 15v5h.553c-1.06-.04-2.105.32-2.991 1H9l-.001 4.386c-.42.194-.724.376-.991.614L8 21H2v-1h6v-5H2v-1h6V9H2V8h6V3zm17.493 17H28v1h-2.376q.444-.498.869-1M15 15H9v5h6zm0-6H9v5h6zm7 0h-6v5h6z"/>' +
      '<path fill="currentColor" fill-opacity="0.28" d="M28 18.13V26h-6.947l.034-.136c.9-.953 2.016-2.14 3.373-3.595l.345-.369A51 51 0 0 0 28 18.13M28 9v.816c-2.48 1.327-5.735 4.029-8.678 7.184-1.037 1.112-1.966 2.096-2.772 3.004-1.256-.053-2.491.466-3.468 1.409-.52.501-.882 1.047-1.318 1.869l-.076.145-.099.189-.098.187c-.194.37-.318.578-.428.716.015-.018-.322.157-.975.416l-.4.158c-.262.105-.489.2-.69.293L9 9z"/>' +
      '<path fill="currentColor" fill-opacity="0.55" d="M23 23.83V26h-1v-1.103zM28 8v1h-5v4.432c-.567.496-1.14 1.02-1.71 1.569L16 15v5h.553c-1.06-.04-2.105.32-2.991 1H9l-.001 4.386c-.42.194-.724.376-.991.614L8 8zm-1.507 12H28v1h-2.376q.444-.498.869-1M15 15H9v5h6zm0-6H9v5h6zm7 0h-6v5h6z"/>' +
      '<g fill-rule="nonzero">' +
      '<path fill="currentColor" fill-opacity="0.42" d="M30.478 10.455c-1.103-1.087-7.663 5.45-10.397 8.221-1.357 1.377-2.717 2.842-3.133 3.393-.181.24.06.314.165.37.538.284.914.546 1.4 1.038.487.492.746.873 1.024 1.42.057.107.13.35.365.166.545-.421 1.812-1.838 3.17-3.212 2.734-2.77 8.508-10.309 7.406-11.396"/>' +
      '<path fill="currentColor" d="M30.81 10.195c1.215 1.236-3.607 7.732-8.068 12.389l-.162.169c-.275.29-.554.587-.91.97.048-.05-.65.704-.824.89-.618.664-.973 1.023-1.223 1.222-.393.314-.76.146-.945-.223l-.046-.096a4.7 4.7 0 0 0-.914-1.306 4.6 4.6 0 0 0-1.255-.959l-.089-.045c-.355-.193-.52-.588-.213-1.008.41-.559 1.711-2.01 3.079-3.438 4.958-5.18 10.38-9.773 11.57-8.565M20.236 23.974c.173-.185.87-.939.823-.889.358-.386.639-.686.918-.979l.163-.172c1.491-1.556 3.746-4.232 5.397-6.455 1.855-2.5 2.881-4.421 2.68-4.627-.212-.214-1.987.875-4.33 2.86a89 89 0 0 0-6.048 5.697c-1.178 1.23-2.313 2.486-2.827 3.121.506.29.869.575 1.304 1.03.442.46.718.843.992 1.371.208-.195.504-.502.928-.957"/>' +
      '<path fill="currentColor" fill-opacity="0.12" d="M14.226 22.778c-.862.776-1.251 2.024-1.89 2.782-.603.715-2.805 1.093-3.403 1.7-.67.68 6.506 3.194 8.83.883 2.26-2.247 2.26-3.904.69-5.365s-3.297-.84-4.227 0"/>' +
      '<path fill="currentColor" fill-opacity="0.55" d="M13.306 24.303q2.504.601 2.937 1.266.44.675.967 2.75a2 2 0 0 0 .162-.157c2.1-2.26 2.246-3.87.725-5.401-1.218-1.228-2.673-1.063-3.766.005-.337.329-.61.744-.967 1.426zm5.417-2.178c1.894 1.908 1.699 4.066-.705 6.652-1.21 1.302-3.528 1.502-6.17.895-1.028-.236-2.016-.585-2.723-.952-.936-.486-1.428-.99-.923-1.545.28-.308.7-.527 1.515-.858l.415-.167c.943-.378 1.386-.61 1.58-.86.179-.229.343-.506.583-.968.027-.05.225-.436.285-.551.398-.759.714-1.24 1.135-1.65 1.412-1.38 3.41-1.606 5.008.004"/>' +
      "</g>" +
      "</g>",
  );
}

/** 单元格样式（32×32，表格 + 格式刷变体，单色层次） */
export function iconCellStyle(): SVGSVGElement {
  return svgElMarkup(
    "0 0 32 32",
    '<g fill="none" fill-rule="evenodd">' +
      '<path fill="currentColor" d="m29 2 .001 7.338q-.474.196-1 .478L28 3H2v23h6.008a3 3 0 0 0-.22.217c-.234.254-.381.519-.457.783H1V2zm0 14.768V27h-8.304q.266-.574.391-1.136l-.034.136H28v-7.87q.528-.69 1-1.362"/>' +
      '<path fill="currentColor" fill-opacity="0.12" d="M28 3v6.816c-2.48 1.327-5.735 4.029-8.678 7.184-1.037 1.112-1.966 2.096-2.772 3.004-1.256-.053-2.491.466-3.468 1.409-.52.501-.882 1.047-1.318 1.869l-.076.145-.099.189-.098.187c-.194.37-.318.578-.428.716.015-.018-.322.157-.975.416l-.4.158c-.811.326-1.285.555-1.68.907H2V3zm0 15.13V26h-6.947l.034-.136c.9-.953 2.016-2.14 3.373-3.595l.345-.369A51 51 0 0 0 28 18.13"/>' +
      '<path fill="currentColor" fill-opacity="0.42" d="M23.999 22.763 24 26h-1v-2.17q.474-.504.999-1.067M7 3v4h16V3h1v4h4v1h-4v4.585q-.495.407-1 .847V8H7v13h6.562a5.5 5.5 0 0 0-1.004 1H7v4H6v-4H2v-1h4V8H2V7h4V3zm18.624 18H28v1h-3.29l.095-.1q.417-.447.82-.9"/>' +
      '<path fill="currentColor" fill-opacity="0.28" d="M23 8v5.432A50 50 0 0 0 19.322 17c-1.037 1.112-1.966 2.096-2.772 3.004-1.058-.045-2.102.316-2.988.995L7 21V8z"/>' +
      '<path fill="currentColor" fill-opacity="0.55" d="M24 7v5.585q-.495.407-1 .847V8H7v13h6.562a5.5 5.5 0 0 0-1.004 1H6V7z"/>' +
      '<g fill-rule="nonzero">' +
      '<path fill="currentColor" fill-opacity="0.42" d="M30.478 10.455c-1.103-1.087-7.663 5.45-10.397 8.221-1.357 1.377-2.717 2.842-3.133 3.393-.181.24.06.314.165.37.538.284.914.546 1.4 1.038.487.492.746.873 1.024 1.42.057.107.13.35.365.166.545-.421 1.812-1.838 3.17-3.212 2.734-2.77 8.508-10.309 7.406-11.396"/>' +
      '<path fill="currentColor" d="M30.81 10.195c1.215 1.236-3.607 7.732-8.068 12.389l-.162.169c-.275.29-.554.587-.91.97.048-.05-.65.704-.824.89-.618.664-.973 1.023-1.223 1.222-.393.314-.76.146-.945-.223l-.046-.096a4.7 4.7 0 0 0-.914-1.306 4.6 4.6 0 0 0-1.255-.959l-.089-.045c-.355-.193-.52-.588-.213-1.008.41-.559 1.711-2.01 3.079-3.438 4.958-5.18 10.38-9.773 11.57-8.565M20.236 23.974c.173-.185.87-.939.823-.889.358-.386.639-.686.918-.979l.163-.172c1.491-1.556 3.746-4.232 5.397-6.455 1.855-2.5 2.881-4.421 2.68-4.627-.212-.214-1.987.875-4.33 2.86a89 89 0 0 0-6.048 5.697c-1.178 1.23-2.313 2.486-2.827 3.121.506.29.869.575 1.304 1.03.442.46.718.843.992 1.371.208-.195.504-.502.928-.957"/>' +
      '<path fill="currentColor" fill-opacity="0.12" d="M14.226 22.778c-.862.776-1.251 2.024-1.89 2.782-.603.715-2.805 1.093-3.403 1.7-.67.68 6.506 3.194 8.83.883 2.26-2.247 2.26-3.904.69-5.365s-3.297-.84-4.227 0"/>' +
      '<path fill="currentColor" fill-opacity="0.55" d="M13.306 24.303q2.504.601 2.937 1.266.44.675.967 2.75a2 2 0 0 0 .162-.157c2.1-2.26 2.246-3.87.725-5.401-1.218-1.228-2.673-1.063-3.766.005-.337.329-.61.744-.967 1.426zm5.417-2.178c1.894 1.908 1.699 4.066-.705 6.652-1.21 1.302-3.528 1.502-6.17.895-1.028-.236-2.016-.585-2.723-.952-.936-.486-1.428-.99-.923-1.545.28-.308.7-.527 1.515-.858l.415-.167c.943-.378 1.386-.61 1.58-.86.179-.229.343-.506.583-.968.027-.05.225-.436.285-.551.398-.759.714-1.24 1.135-1.65 1.412-1.38 3.41-1.606 5.008.004"/>' +
      "</g>" +
      "</g>",
  );
}

/** Ribbon 下拉小三角（8×8，与资源几何一致） */
export function iconChevronDown(): SVGSVGElement {
  return svgElMarkup(
    "0 0 8 8",
    '<path fill="currentColor" fill-rule="evenodd" d="M2 3h5L4.5 6z"/>',
  );
}
