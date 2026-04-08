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
    '<g fill="none" fill-rule="nonzero">' +
      '<path fill="currentColor" fill-opacity="0.34" d="M0 12h16v4H0z"/>' +
      '<path fill="currentColor" d="M1 13v2h14v-2zm-1-1h16v4H0z"/>' +
      '<path fill="currentColor" fill-opacity="0.3" d="M6.993 3.33 2.331 7.994l3.676 3.676 4.662-4.662z"/>' +
      '<path fill="currentColor" d="M6.644 3.68 2.33 7.994l3.676 3.676 4.662-4.662L7.592 3.93v2.525l.006.044L8.1 7l-1 1.002L6.1 7l.544-.544zm0-1.33v-.84c0-.323-.204-.547-.493-.553H6.06c-.28 0-.511.24-.511.553v1.935zM4.6 4.393V1.51c0-.324.1-.636.282-.893q.145-.205.346-.35C5.47.094 5.758 0 6.06 0h.073c.255 0 .5.067.715.192.458.266.744.769.744 1.318V2.6L12 7.007 6.007 13 1 7.993z"/>' +
      '<path fill="currentColor" fill-opacity="0.48" d="M11.739 9.267c0 1.181.248 2.142.638 2.733 0-.628.14-1.035.53-1.773.922-1.514 1.594-2.733.603-3.988Q12.518 4.983 10 5s1.393 1.595 1.739 1.977c.142.628 0 1.108 0 2.29"/>' +
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

export function iconPercent(): SVGSVGElement {
  return svgEl("0 0 24 24", '<line x1="19" y1="5" x2="5" y2="19"/><circle cx="6.5" cy="6.5" r="2.5"/><circle cx="17.5" cy="17.5" r="2.5"/>');
}

export function iconCommaStyle(): SVGSVGElement {
  return svgEl("0 0 24 24", '<path d="M7 8h10"/><path d="M7 12h6"/><path d="M12 16v4"/><path d="M10 20h4"/>');
}

export function iconChart(): SVGSVGElement {
  return svgEl("0 0 24 24", '<line x1="18" y1="20" x2="18" y2="10"/><line x1="12" y1="20" x2="12" y2="4"/><line x1="6" y1="20" x2="6" y2="14"/>');
}

export function iconTable(): SVGSVGElement {
  return svgEl("0 0 24 24", '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18"/><path d="M3 15h18"/><path d="M9 3v18"/>');
}

export function iconImage(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<rect x="3" y="3" width="18" height="18" rx="2"/><circle cx="8.5" cy="8.5" r="1.5"/><path d="M21 15l-5-5L5 21"/>',
  );
}

export function iconShapes(): SVGSVGElement {
  return svgEl("0 0 24 24", '<circle cx="8" cy="8" r="3"/><rect x="13" y="5" width="6" height="6" rx="1"/><path d="M3 21l7-7 4 4 7-7"/>');
}

export function iconSparkline(): SVGSVGElement {
  return svgEl("0 0 24 24", '<path d="M3 17l6-6 4 4 8-8"/><path d="M14 7h7v7"/>');
}

export function iconOrientation(): SVGSVGElement {
  return svgEl("0 0 24 24", '<rect x="4" y="4" width="16" height="10" rx="1"/><path d="M8 20h8"/><path d="M12 14v6"/>');
}

export function iconMargins(): SVGSVGElement {
  return svgEl("0 0 24 24", '<rect x="6" y="6" width="12" height="12"/><path d="M3 3v18"/><path d="M21 3v18"/><path d="M3 3h18"/><path d="M3 21h18"/>');
}

export function iconBackground(): SVGSVGElement {
  return svgEl("0 0 24 24", '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 15l6-6 4 4 8-8"/>');
}

export function iconPrintTitles(): SVGSVGElement {
  return svgEl("0 0 24 24", '<rect x="4" y="4" width="16" height="16" rx="1"/><path d="M4 8h16"/><path d="M8 4v4"/>');
}

export function iconFunction(): SVGSVGElement {
  return svgEl("0 0 24 24", '<path d="M4 19c0-4 2-6 6-6s6 2 6 6"/><path d="M10 13V9a2 2 0 1 1 4 0v4"/>');
}

export function iconAudit(): SVGSVGElement {
  return svgEl("0 0 24 24", '<circle cx="12" cy="12" r="10"/><path d="M12 6v6l4 2"/>');
}

export function iconSort(): SVGSVGElement {
  return svgEl("0 0 24 24", '<path d="M3 6h7"/><path d="M3 12h5"/><path d="M3 18h3"/><path d="M16 5v14l4-4"/>');
}

export function iconFilter(): SVGSVGElement {
  return svgEl("0 0 24 24", '<path d="M4 4h16l-6 8v6l-4 2v-8z"/>');
}

export function iconSlicer(): SVGSVGElement {
  return svgEl("0 0 24 24", '<rect x="4" y="4" width="16" height="6" rx="1"/><rect x="4" y="14" width="10" height="6" rx="1"/><circle cx="17" cy="17" r="2"/>');
}

export function iconDataTools(): SVGSVGElement {
  return svgEl("0 0 24 24", '<ellipse cx="12" cy="5" rx="9" ry="3"/><path d="M21 12c0 1.66-4 3-9 3s-9-1.34-9-3"/><path d="M3 5v14c0 1.66 4 3 9 3s9-1.34 9-3V5"/>');
}

export function iconNormalView(): SVGSVGElement {
  return svgEl("0 0 24 24", '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18"/>');
}

export function iconPageBreak(): SVGSVGElement {
  return svgEl("0 0 24 24", '<rect x="3" y="3" width="18" height="18" rx="1"/><path d="M12 3v18"/><path d="M3 12h18" stroke-dasharray="2 2"/>');
}

export function iconGrid(): SVGSVGElement {
  return svgEl("0 0 24 24", '<rect x="3" y="3" width="18" height="18"/><path d="M3 9h18M3 15h18M9 3v18M15 3v18"/>');
}

export function iconZoom(): SVGSVGElement {
  return svgEl("0 0 24 24", '<circle cx="11" cy="11" r="8"/><path d="M21 21l-4.35-4.35"/><path d="M11 8v6M8 11h6"/>');
}

export function iconFreeze(): SVGSVGElement {
  return svgEl("0 0 24 24", '<rect x="3" y="3" width="18" height="18"/><path d="M9 3v18"/><path d="M3 9h18"/>');
}

export function iconWindow(): SVGSVGElement {
  return svgEl("0 0 24 24", '<rect x="3" y="3" width="8" height="8" rx="1"/><rect x="13" y="13" width="8" height="8" rx="1"/><path d="M13 3h8v8"/>');
}

export function iconRuler(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<path d="M4 20h16"/><path d="M5 20v-5M9 20v-3M13 20v-5M17 20v-3"/>',
  );
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

/** Ribbon 下拉小三角（8×8，与资源几何一致） */
export function iconChevronDown(): SVGSVGElement {
  return svgElMarkup("0 0 8 8", '<path fill="currentColor" fill-rule="evenodd" d="M2 3h5L4.5 6z"/>');
}
