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

export function iconPaste(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<path d="M9 4h6a2 2 0 0 1 2 2v14a2 2 0 0 1-2 2H9a2 2 0 0 1-2-2V6a2 2 0 0 1 2-2z"/><path d="M8 4V3a1 1 0 0 1 1-1h6a1 1 0 0 1 1 1v1"/>',
  );
}

export function iconCut(): SVGSVGElement {
  return svgEl("0 0 24 24", '<circle cx="6" cy="6" r="3"/><circle cx="6" cy="18" r="3"/><path d="M20 4L8.12 15.88"/><path d="M14.47 14.48L20 20"/><path d="M8.12 8.12L12 12"/>');
}

export function iconCopy(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<rect x="9" y="9" width="13" height="13" rx="2"/><path d="M5 15H4a2 2 0 0 1-2-2V4a2 2 0 0 1 2-2h9a2 2 0 0 1 2 2v1"/>',
  );
}

export function iconFormatPainter(): SVGSVGElement {
  return svgEl("0 0 24 24", '<path d="M12 19l7-7 3 3-7 7-3-3z"/><path d="M18 13l-1.5-7.5L2 2l3.5 14.5L13 18l5-5z"/><path d="M2 2l7.586 7.586"/><circle cx="11" cy="11" r="2"/>');
}

export function iconBold(): SVGSVGElement {
  return svgEl("0 0 24 24", '<path d="M6 4h8a4 4 0 0 1 0 8H6z"/><path d="M6 12h9a4 4 0 0 1 0 8H6z"/>');
}

export function iconItalic(): SVGSVGElement {
  return svgEl("0 0 24 24", '<line x1="19" y1="4" x2="10" y2="4"/><line x1="14" y1="20" x2="5" y2="20"/><line x1="15" y1="4" x2="9" y2="20"/>');
}

export function iconUnderline(): SVGSVGElement {
  return svgEl("0 0 24 24", '<path d="M6 3v7a6 6 0 0 0 12 0V3"/><line x1="4" y1="21" x2="20" y2="21"/>');
}

export function iconFillColor(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<path d="M14 4l6 6-8 8-6-6z"/><path d="M3 21h12"/><path d="M17 7l2-2"/>',
  );
}

export function iconFontColor(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<path d="M6 18l6-14 6 14"/><path d="M9 12h6"/><path d="M4 21h16"/>',
  );
}

export function iconAlignLeft(): SVGSVGElement {
  return svgEl("0 0 24 24", '<line x1="3" y1="6" x2="21" y2="6"/><line x1="3" y1="12" x2="15" y2="12"/><line x1="3" y1="18" x2="18" y2="18"/>');
}

export function iconAlignCenter(): SVGSVGElement {
  return svgEl("0 0 24 24", '<line x1="3" y1="6" x2="21" y2="6"/><line x1="6" y1="12" x2="18" y2="12"/><line x1="5" y1="18" x2="19" y2="18"/>');
}

export function iconAlignRight(): SVGSVGElement {
  return svgEl("0 0 24 24", '<line x1="3" y1="6" x2="21" y2="6"/><line x1="9" y1="12" x2="21" y2="12"/><line x1="6" y1="18" x2="21" y2="18"/>');
}

export function iconAlignTop(): SVGSVGElement {
  return svgEl("0 0 24 24", '<line x1="6" y1="4" x2="6" y2="20"/><line x1="12" y1="4" x2="12" y2="16"/><line x1="18" y1="4" x2="18" y2="12"/>');
}

export function iconAlignMiddle(): SVGSVGElement {
  return svgEl("0 0 24 24", '<line x1="6" y1="4" x2="6" y2="20"/><rect x="9" y="9" width="6" height="6" rx="1"/><line x1="18" y1="4" x2="18" y2="20"/>');
}

export function iconAlignBottom(): SVGSVGElement {
  return svgEl("0 0 24 24", '<line x1="6" y1="4" x2="6" y2="20"/><line x1="12" y1="8" x2="12" y2="20"/><line x1="18" y1="12" x2="18" y2="20"/>');
}

export function iconWrapText(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<path d="M3 6h18"/><path d="M3 12h10"/><path d="M3 18h18"/><path d="M17 8v8"/><path d="M14 11l3 3 3-3"/>',
  );
}

export function iconMerge(): SVGSVGElement {
  return svgEl("0 0 24 24", '<rect x="3" y="3" width="18" height="18" rx="2"/><path d="M3 9h18"/><path d="M9 3v18"/>');
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

/** 逆时针弯箭头（撤销） */
export function iconUndo(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<path d="M9 9H5l3-3"/><path d="M5 9l3 3"/><path d="M9 9h6a5 5 0 1 1 0 10h-7"/>',
  );
}

/** 顺时针弯箭头（重做） */
export function iconRedo(): SVGSVGElement {
  return svgEl(
    "0 0 24 24",
    '<path d="M15 9h4l-3-3"/><path d="M19 9l-3 3"/><path d="M15 9H9a5 5 0 1 0 0 10h7"/>',
  );
}

export function iconChevronDown(): SVGSVGElement {
  return svgEl("0 0 24 24", '<path d="M6 9l6 6 6-6"/>');
}
