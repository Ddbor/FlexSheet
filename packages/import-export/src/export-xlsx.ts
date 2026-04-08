import {
  Cell,
  type CellBorderKind,
  type CellBorderSide,
  type CellScalar,
  type CellStyle,
} from "@flexsheet/core";
import type { Workbook } from "@flexsheet/core";
import type { Worksheet } from "@flexsheet/core";
import { formatCellRef } from "./a1.js";
import { escapeXml, sanitizeXml10Text } from "./xml-escape.js";
import { buildZipArchive, type ZipEntryInput } from "./zip-writer.js";

const SS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const CP_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties";
const DC_NS = "http://purl.org/dc/elements/1.1/";
const DCTERMS_NS = "http://purl.org/dc/terms/";
const XSI_NS = "http://www.w3.org/2001/XMLSchema-instance";

function ooxmlFormula(raw: string): string {
  let t = raw.trim();
  if (t.startsWith("=")) {
    t = t.slice(1);
  }
  return sanitizeXml10Text(t);
}

function borderSideSig(side: CellBorderSide | undefined): string {
  if (side === undefined) {
    return "";
  }
  return `${side.kind}|${side.colorArgb ?? ""}`;
}

function firstFontNameFromStack(stack: string | undefined): string {
  if (stack === undefined || stack.trim() === "") {
    return "";
  }
  const first =
    stack
      .split(",")[0]
      ?.trim()
      .replace(/^["']|["']$/g, "") ?? "";
  return first;
}

/**
 * 与画布 `resolvedVAlign` 一致：未设置或非 top/bottom 时视为 middle。
 * 导出时必须写入 OOXML `vertical`，否则 Excel 按默认底部对齐，与 FlexSheet 显示不一致。
 */
function resolveVAlignForExport(v: string | undefined): "top" | "middle" | "bottom" {
  if (v === "top" || v === "bottom") {
    return v;
  }
  return "middle";
}

function styleSignature(st: CellStyle | null | undefined): string {
  if (st === null || st === undefined) {
    return "";
  }
  const tor = st.textOrientation;
  return JSON.stringify({
    b: st.bold === true,
    i: st.italic === true,
    fg: st.fgArgb ?? "",
    fill: st.fillArgb ?? "",
    ff: firstFontNameFromStack(st.fontFamily),
    fs: st.fontSizePt ?? 0,
    ul: st.underline ?? "",
    ha: st.hAlign ?? "",
    va: resolveVAlignForExport(st.vAlign),
    ind: st.indentLevel ?? 0,
    wx: st.wrapText === true,
    tor: tor !== undefined && tor !== "horizontal" ? tor : "",
    bt: borderSideSig(st.borderTop),
    bl: borderSideSig(st.borderLeft),
    bb: borderSideSig(st.borderBottom),
    br: borderSideSig(st.borderRight),
  });
}

interface StyleTable {
  readonly xfBySig: Map<string, number>;
  readonly fontsXml: string[];
  readonly fillsXml: string[];
  readonly bordersXml: string[];
  readonly cellXfsXml: string[];
}

/** `styleSignature` 与 `ensureStyle` 中 JSON 往返用的稳定形状。 */
interface StyleSignaturePayload {
  readonly b: boolean;
  readonly i?: boolean;
  readonly fg: string;
  readonly fill: string;
  readonly ff?: string;
  /** 0 表示未设置字号（与默认 11pt 一致）。 */
  readonly fs?: number;
  readonly ul?: string;
  readonly ha: string;
  readonly va: string;
  readonly ind?: number;
  readonly wx?: boolean;
  /** 非空时为 `CellTextOrientation`（不含 horizontal）。 */
  readonly tor?: string;
  readonly bt?: string;
  readonly bl?: string;
  readonly bb?: string;
  readonly br?: string;
}

const DEFAULT_FONT_NAME = "Calibri";
const DEFAULT_FONT_SIZE_PT = 11;

function minimalStyleTable(): StyleTable {
  const xfBySig = new Map<string, number>();
  xfBySig.set("", 0);
  const fontsXml: string[] = [
    `<font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font>`,
  ];
  const fillsXml: string[] = [
    `<fill><patternFill patternType="none"/></fill>`,
    `<fill><patternFill patternType="gray125"/></fill>`,
  ];
  const bordersXml: string[] = [`<border><left/><right/><top/><bottom/><diagonal/></border>`];
  const cellXfsXml: string[] = [`<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`];
  return { xfBySig, fontsXml, fillsXml, bordersXml, cellXfsXml };
}

function ooxmlBorderStyle(kind: CellBorderKind): string {
  return kind;
}

function parseBorderSideToken(tok: string): CellBorderSide | undefined {
  if (tok === "") {
    return undefined;
  }
  const bar = tok.indexOf("|");
  const kind = (bar >= 0 ? tok.slice(0, bar) : tok) as CellBorderKind;
  const color = bar >= 0 ? tok.slice(bar + 1) : "";
  if (kind === "") {
    return undefined;
  }
  return color.length > 0 ? { kind, colorArgb: color } : { kind };
}

function borderElementFromSignature(st: StyleSignaturePayload): string {
  const L = parseBorderSideToken(st.bl ?? "");
  const R = parseBorderSideToken(st.br ?? "");
  const T = parseBorderSideToken(st.bt ?? "");
  const B = parseBorderSideToken(st.bb ?? "");
  const side = (
    s: CellBorderSide | undefined,
    tag: "left" | "right" | "top" | "bottom",
  ): string => {
    if (s === undefined) {
      return `<${tag}/>`;
    }
    const rgb =
      s.colorArgb !== undefined && s.colorArgb !== "" ? escapeXml(s.colorArgb) : "FF000000";
    return `<${tag} style="${ooxmlBorderStyle(s.kind)}"><color rgb="${rgb}"/></${tag}>`;
  };
  return `<border>${side(L, "left")}${side(R, "right")}${side(T, "top")}${side(B, "bottom")}<diagonal/></border>`;
}

function borderSignatureKey(st: StyleSignaturePayload): string {
  return `${st.bl ?? ""};${st.br ?? ""};${st.bt ?? ""};${st.bb ?? ""}`;
}

/** OOXML `alignment/@textRotation`（与 Excel 常用预设一致）。 */
function ooxmlTextRotationFromTor(tor: string): number | undefined {
  switch (tor) {
    case "angleUp45":
      return 45;
    case "angleDown45":
      return 135;
    case "verticalStack":
      return 255;
    case "rotateUp90":
      return 90;
    case "rotateDown90":
      return 180;
    default:
      return undefined;
  }
}

function buildStyleTable(workbook: Workbook, opts: XlsxExportOptions): StyleTable {
  if (!opts.includeStyles) {
    return minimalStyleTable();
  }

  const xfBySig = new Map<string, number>();
  xfBySig.set("", 0);

  const fontsXml: string[] = [
    `<font><sz val="11"/><color rgb="FF000000"/><name val="Calibri"/><family val="2"/></font>`,
  ];
  const fillsXml: string[] = [
    `<fill><patternFill patternType="none"/></fill>`,
    `<fill><patternFill patternType="gray125"/></fill>`,
  ];
  const bordersXml: string[] = [`<border><left/><right/><top/><bottom/><diagonal/></border>`];
  const cellXfsXml: string[] = [`<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`];

  let nextFont = 1;
  let nextFill = 2;
  let nextBorder = 1;
  let nextXf = 1;

  const borderByKey = new Map<string, number>();

  const ensureStyle = (sig: string): void => {
    if (sig === "" || xfBySig.has(sig)) {
      return;
    }
    const st = JSON.parse(sig) as StyleSignaturePayload;
    const fontName = st.ff !== undefined && st.ff !== "" ? st.ff : DEFAULT_FONT_NAME;
    const fontSize = st.fs !== undefined && st.fs > 0 ? st.fs : DEFAULT_FONT_SIZE_PT;
    const needFont =
      st.b === true ||
      st.i === true ||
      st.fg !== "" ||
      (st.ff !== undefined && st.ff !== "" && st.ff !== DEFAULT_FONT_NAME) ||
      (st.fs !== undefined && st.fs > 0 && st.fs !== DEFAULT_FONT_SIZE_PT) ||
      (st.ul !== undefined && st.ul !== "");

    let fontId = 0;
    if (needFont) {
      const bold = st.b ? "<b/>" : "";
      const italic = st.i ? "<i/>" : "";
      let underline = "";
      if (st.ul === "double") {
        underline = `<u val="double"/>`;
      } else if (st.ul === "single") {
        underline = `<u val="single"/>`;
      }
      const color = st.fg !== "" ? `<color rgb="${escapeXml(st.fg)}"/>` : `<color rgb="FF000000"/>`;
      const sz = Math.max(1, Math.min(409, Math.round(fontSize * 100) / 100));
      const szAttr = Number.isInteger(sz) ? String(sz) : sz.toFixed(2).replace(/\.?0+$/, "");
      fontsXml.push(
        `<font>${bold}${italic}${underline}<sz val="${szAttr}"/>${color}<name val="${escapeXml(fontName)}"/><family val="2"/></font>`,
      );
      fontId = nextFont++;
    }

    let fillId = 0;
    if (st.fill !== "") {
      fillsXml.push(
        `<fill><patternFill patternType="solid"><fgColor rgb="${escapeXml(st.fill)}"/><bgColor indexed="64"/></patternFill></fill>`,
      );
      fillId = nextFill++;
    }

    const bKey = borderSignatureKey(st);
    let borderId = 0;
    if (bKey !== ";;;") {
      const existing = borderByKey.get(bKey);
      if (existing !== undefined) {
        borderId = existing;
      } else {
        borderId = nextBorder++;
        bordersXml.push(borderElementFromSignature(st));
        borderByKey.set(bKey, borderId);
      }
    }

    const applyFont = needFont ? ` applyFont="1"` : "";
    const applyFill = st.fill !== "" ? ` applyFill="1"` : "";
    const applyBorder = borderId > 0 ? ` applyBorder="1"` : "";
    const alignAttrs: string[] = [];
    /**
     * OOXML 的 `indent` 仅在同时声明 `horizontal` 时 Excel 才会按预期显示（左对齐为左侧缩进，右对齐为右侧缩进）。
     * 未写 horizontal 时等价于常规对齐，Excel 会忽略 indent；与画布默认 `resolvedHAlign` 为 left 一致。
     */
    let ha = st.ha;
    if (ha !== "left" && ha !== "center" && ha !== "right") {
      ha = "";
    }
    const ind = st.ind ?? 0;
    if (ind > 0 && ha !== "center" && ha !== "right") {
      ha = "left";
    }
    if (ha === "left" || ha === "center" || ha === "right") {
      alignAttrs.push(`horizontal="${ha}"`);
    }
    const vAlign = resolveVAlignForExport(st.va);
    alignAttrs.push(`vertical="${vAlign === "middle" ? "center" : vAlign}"`);
    if (ind > 0) {
      alignAttrs.push(`indent="${Math.min(255, Math.round(ind))}"`);
    }
    if (st.wx === true) {
      alignAttrs.push(`wrapText="1"`);
    }
    const tor = st.tor ?? "";
    if (tor !== "") {
      const tr = ooxmlTextRotationFromTor(tor);
      if (tr !== undefined) {
        alignAttrs.push(`textRotation="${tr}"`);
      }
    }
    const alignInner = alignAttrs.length > 0 ? `<alignment ${alignAttrs.join(" ")}/>` : "";
    const applyAlignment = alignInner !== "" ? ` applyAlignment="1"` : "";
    if (alignInner === "") {
      cellXfsXml.push(
        `<xf numFmtId="0" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${applyFont}${applyFill}${applyBorder}/>`,
      );
    } else {
      cellXfsXml.push(
        `<xf numFmtId="0" fontId="${fontId}" fillId="${fillId}" borderId="${borderId}" xfId="0"${applyFont}${applyFill}${applyBorder}${applyAlignment}>${alignInner}</xf>`,
      );
    }
    xfBySig.set(sig, nextXf++);
  };

  for (let i = 0; i < workbook.sheetCount; i++) {
    const sh = workbook.getSheet(i);
    if (sh === undefined) {
      continue;
    }
    sh.iterateCells((c) => {
      if (!shouldExportCellForXlsx(sh, c, opts)) {
        return;
      }
      ensureStyle(styleSignature(effectiveCellStyleForXlsxExport(sh, c, opts)));
    });
    for (const reg of sh.getMergeRegionsSnapshot()) {
      if (reg.rowSpan <= 1 && reg.colSpan <= 1) {
        continue;
      }
      const endR = reg.masterRow + reg.rowSpan - 1;
      const endC = reg.masterCol + reg.colSpan - 1;
      for (let r = reg.masterRow; r <= endR; r++) {
        for (let c = reg.masterCol; c <= endC; c++) {
          if (r === reg.masterRow && c === reg.masterCol) {
            continue;
          }
          const tmp = new Cell(r, c, null);
          const eff = effectiveCellStyleForXlsxExport(sh, tmp, opts);
          if (eff === null || Object.keys(eff).length === 0) {
            continue;
          }
          ensureStyle(styleSignature(eff));
        }
      }
    }
  }

  return { xfBySig, fontsXml, fillsXml, bordersXml, cellXfsXml };
}

function buildStylesXml(table: StyleTable): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<styleSheet xmlns="${SS_MAIN}">` +
    `<numFmts count="0"/>` +
    `<fonts count="${table.fontsXml.length}">${table.fontsXml.join("")}</fonts>` +
    `<fills count="${table.fillsXml.length}">${table.fillsXml.join("")}</fills>` +
    `<borders count="${table.bordersXml.length}">${table.bordersXml.join("")}</borders>` +
    `<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>` +
    `<cellXfs count="${table.cellXfsXml.length}">${table.cellXfsXml.join("")}</cellXfs>` +
    `<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>` +
    `</styleSheet>`
  );
}

/** XLSX 导出选项（与 Backstage「导出 Excel」复选框对应；未实现项仅占位）。 */
export interface XlsxExportOptions {
  readonly includeStyles: boolean;
  readonly includeFormulas: boolean;
  /** 导出仅有样式、无公式且无值的单元格。 */
  readonly includeSparseStyledEmpty: boolean;
}

export const DEFAULT_XLSX_EXPORT_OPTIONS: XlsxExportOptions = {
  includeStyles: true,
  includeFormulas: true,
  includeSparseStyledEmpty: true,
};

/** 从样式中去掉四边边框，供合并格按边拆分后与主格非边框属性合并。 */
function stripBorderSides(st: CellStyle | null | undefined): CellStyle {
  if (st === null || st === undefined) {
    return {};
  }
  const { borderTop: _t, borderLeft: _l, borderBottom: _b, borderRight: _r, ...rest } = st;
  return rest;
}

/**
 * Excel 合并区只把主格写入 sheetData；边框按单元格几何绘制。
 * 将主格上的四边样式映射到合并矩形周界上每个单元格应有的边（与 Excel 外边框一致）。
 */
function mergeRectBorderSidesAt(
  borderSource: CellStyle | null | undefined,
  row: number,
  col: number,
  masterRow: number,
  masterCol: number,
  endRow: number,
  endCol: number,
): CellStyle {
  if (borderSource === null || borderSource === undefined) {
    return {};
  }
  const onTop = row === masterRow;
  const onBottom = row === endRow;
  const onLeft = col === masterCol;
  const onRight = col === endCol;
  if (!onTop && !onBottom && !onLeft && !onRight) {
    return {};
  }
  const out: CellStyle = {};
  if (onTop && borderSource.borderTop !== undefined) {
    out.borderTop = borderSource.borderTop;
  }
  if (onBottom && borderSource.borderBottom !== undefined) {
    out.borderBottom = borderSource.borderBottom;
  }
  if (onLeft && borderSource.borderLeft !== undefined) {
    out.borderLeft = borderSource.borderLeft;
  }
  if (onRight && borderSource.borderRight !== undefined) {
    out.borderRight = borderSource.borderRight;
  }
  return out;
}

/**
 * 导出用有效样式：多格合并时边框取自主格并按周界拆分；其余属性仍用当前格（通常即主格）。
 */
function effectiveCellStyleForXlsxExport(
  sheet: Worksheet,
  cell: Cell,
  opts: XlsxExportOptions,
): CellStyle | null {
  if (!opts.includeStyles) {
    return cell.style;
  }
  const info = sheet.getMergedRectInfo(cell.row, cell.col);
  if (info.rowSpan <= 1 && info.colSpan <= 1) {
    return cell.style;
  }
  const endR = info.anchorRow + info.rowSpan - 1;
  const endC = info.anchorCol + info.colSpan - 1;
  const master = sheet.getCell(info.anchorRow, info.anchorCol);
  const borderParts = mergeRectBorderSidesAt(
    master.style,
    cell.row,
    cell.col,
    info.anchorRow,
    info.anchorCol,
    endR,
    endC,
  );
  const rest = stripBorderSides(cell.style);
  const merged: CellStyle = { ...rest, ...borderParts };
  return Object.keys(merged).length > 0 ? merged : null;
}

function shouldExportCellForXlsx(sheet: Worksheet, cell: Cell, opts: XlsxExportOptions): boolean {
  if (sheet.isMergeCoveredCell(cell.row, cell.col)) {
    return false;
  }
  const hasF = cell.formula !== null && cell.formula.length > 0;
  const hasV = cell.value !== null && cell.value !== "";
  const hasSt = cell.style !== null && Object.keys(cell.style).length > 0;
  if (hasF && opts.includeFormulas) {
    return true;
  }
  if (hasF && !opts.includeFormulas) {
    return hasV;
  }
  if (hasV) {
    return true;
  }
  if (opts.includeSparseStyledEmpty && hasSt && opts.includeStyles) {
    return true;
  }
  return false;
}

function usedBoundsForSheet(
  sheet: Worksheet,
  opts: XlsxExportOptions,
): {
  minR: number;
  maxR: number;
  minC: number;
  maxC: number;
} | null {
  let minR = Infinity;
  let maxR = -1;
  let minC = Infinity;
  let maxC = -1;
  let hasCell = false;
  sheet.iterateCells((c) => {
    if (!shouldExportCellForXlsx(sheet, c, opts)) {
      return;
    }
    hasCell = true;
    minR = Math.min(minR, c.row);
    maxR = Math.max(maxR, c.row);
    minC = Math.min(minC, c.col);
    maxC = Math.max(maxC, c.col);
  });
  for (const r of sheet.getMergeRegionsSnapshot()) {
    if (r.rowSpan <= 1 && r.colSpan <= 1) {
      continue;
    }
    const endR = r.masterRow + r.rowSpan - 1;
    const endC = r.masterCol + r.colSpan - 1;
    hasCell = true;
    minR = Math.min(minR, r.masterRow);
    maxR = Math.max(maxR, endR);
    minC = Math.min(minC, r.masterCol);
    maxC = Math.max(maxC, endC);
  }
  if (!hasCell) {
    return null;
  }
  return { minR, maxR, minC, maxC };
}

/** FlexSheet 行高为逻辑 px（与 import 互逆），OOXML `ht` 为磅。 */
function rowHeightPxToOoxmlHt(px: number): string {
  const pt = (px * 72) / 96;
  return pt.toFixed(2);
}

/** 与 `import-xlsx` 互逆：列宽 px → OOXML `width`（字符宽度近似）。 */
function colWidthPxToOoxmlWidth(px: number): string {
  const w = Math.max(0, (px - 5) / 7);
  return w.toFixed(2);
}

/** 合并连续同宽列，减少 `<col>` 数量。 */
function buildColsXml(sheet: Worksheet): string {
  const n = Math.max(1, sheet.colCount);
  const parts: string[] = [];
  let c = 0;
  while (c < n) {
    const wPx = sheet.getColWidth(c);
    const wStr = colWidthPxToOoxmlWidth(wPx);
    let cEnd = c;
    while (cEnd + 1 < n && sheet.getColWidth(cEnd + 1) === wPx) {
      cEnd++;
    }
    const min = c + 1;
    const max = cEnd + 1;
    const custom = wPx !== sheet.defaultColWidth ? ` customWidth="1"` : "";
    parts.push(`<col min="${min}" max="${max}" width="${wStr}"${custom}/>`);
    c = cEnd + 1;
  }
  return `<cols>${parts.join("")}</cols>`;
}

function rowXmlAttributesForHeight(sheet: Worksheet, rowIndex: number): string {
  const hPx = sheet.getRowHeight(rowIndex);
  if (hPx === sheet.defaultRowHeight) {
    return "";
  }
  return ` ht="${rowHeightPxToOoxmlHt(hPx)}" customHeight="1"`;
}

function cachedValueXml(value: CellScalar): string {
  if (value === null || value === "") {
    return "";
  }
  if (typeof value === "number") {
    return `<v>${value}</v>`;
  }
  if (typeof value === "boolean") {
    return `<v>${value ? 1 : 0}</v>`;
  }
  return `<v>${escapeXml(String(value))}</v>`;
}

function cellToXml(
  sheet: Worksheet,
  cell: Cell,
  sst: Map<string, number>,
  xfBySig: Map<string, number>,
  opts: XlsxExportOptions,
): string {
  const ref = formatCellRef(cell.row, cell.col);
  const sig = opts.includeStyles
    ? styleSignature(effectiveCellStyleForXlsxExport(sheet, cell, opts))
    : "";
  const xf = opts.includeStyles ? (xfBySig.get(sig) ?? 0) : 0;
  const sAttr = opts.includeStyles && xf > 0 ? ` s="${xf}"` : "";

  const hasF = cell.formula !== null && cell.formula.length > 0;
  if (hasF && opts.includeFormulas) {
    const f = escapeXml(ooxmlFormula(cell.formula as string));
    const vPart = cachedValueXml(cell.value);
    const tStr = typeof cell.value === "string" && cell.value !== "" ? ` t="str"` : "";
    return `<c r="${ref}"${sAttr}${tStr}><f>${f}</f>${vPart}</c>`;
  }

  if (cell.value === null || cell.value === "") {
    return `<c r="${ref}"${sAttr}/>`;
  }
  if (typeof cell.value === "number") {
    return `<c r="${ref}"${sAttr}><v>${cell.value}</v></c>`;
  }
  if (typeof cell.value === "boolean") {
    return `<c r="${ref}"${sAttr} t="b"><v>${cell.value ? 1 : 0}</v></c>`;
  }
  const si = sst.get(String(cell.value)) ?? 0;
  return `<c r="${ref}"${sAttr} t="s"><v>${si}</v></c>`;
}

function buildSharedStrings(
  workbook: Workbook,
  opts: XlsxExportOptions,
): { xml: string; index: Map<string, number> } {
  const ordered: string[] = [];
  const index = new Map<string, number>();
  let stringRefCount = 0;
  const add = (s: string): void => {
    if (index.has(s)) {
      return;
    }
    index.set(s, ordered.length);
    ordered.push(s);
  };

  for (let i = 0; i < workbook.sheetCount; i++) {
    const sh = workbook.getSheet(i);
    if (sh === undefined) {
      continue;
    }
    sh.iterateCells((c) => {
      if (!shouldExportCellForXlsx(sh, c, opts)) {
        return;
      }
      const asLiteral = c.formula === null || !opts.includeFormulas;
      if (asLiteral && typeof c.value === "string" && c.value !== "") {
        add(c.value);
        stringRefCount += 1;
      }
    });
  }

  const parts = ordered.map((s) => {
    const body = escapeXml(s);
    const preserve =
      s !== s.trim() || s.includes("\n") || s.includes("\r") ? ` xml:space="preserve"` : "";
    return `<si><t${preserve}>${body}</t></si>`;
  });
  const xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<sst xmlns="${SS_MAIN}" count="${stringRefCount}" uniqueCount="${ordered.length}">` +
    parts.join("") +
    `</sst>`;
  return { xml, index };
}

function mergeCellsXml(sheet: Worksheet): string {
  const regions = sheet.getMergeRegionsSnapshot();
  const parts: string[] = [];
  for (const r of regions) {
    if (r.rowSpan <= 1 && r.colSpan <= 1) {
      continue;
    }
    const endR = r.masterRow + r.rowSpan - 1;
    const endC = r.masterCol + r.colSpan - 1;
    const ref = `${formatCellRef(r.masterRow, r.masterCol)}:${formatCellRef(endR, endC)}`;
    parts.push(`<mergeCell ref="${ref}"/>`);
  }
  if (parts.length === 0) {
    return "";
  }
  return `<mergeCells count="${parts.length}">${parts.join("")}</mergeCells>`;
}

function buildSheetXml(
  sheet: Worksheet,
  sheetIndex: number,
  sst: Map<string, number>,
  xfBySig: Map<string, number>,
  opts: XlsxExportOptions,
): string {
  const b = usedBoundsForSheet(sheet, opts);
  const dim =
    b === null ? "A1" : `${formatCellRef(b.minR, b.minC)}:${formatCellRef(b.maxR, b.maxC)}`;

  const byRow = new Map<number, Cell[]>();
  sheet.iterateCells((c) => {
    if (!shouldExportCellForXlsx(sheet, c, opts)) {
      return;
    }
    const arr = byRow.get(c.row);
    if (arr === undefined) {
      byRow.set(c.row, [c]);
    } else {
      arr.push(c);
    }
  });
  for (const reg of sheet.getMergeRegionsSnapshot()) {
    if (reg.rowSpan <= 1 && reg.colSpan <= 1) {
      continue;
    }
    const endR = reg.masterRow + reg.rowSpan - 1;
    const endC = reg.masterCol + reg.colSpan - 1;
    for (let r = reg.masterRow; r <= endR; r++) {
      for (let c = reg.masterCol; c <= endC; c++) {
        if (r === reg.masterRow && c === reg.masterCol) {
          continue;
        }
        const tmp = new Cell(r, c, null);
        const eff = effectiveCellStyleForXlsxExport(sheet, tmp, opts);
        if (eff === null || Object.keys(eff).length === 0) {
          continue;
        }
        const arr = byRow.get(r);
        if (arr === undefined) {
          byRow.set(r, [tmp]);
        } else {
          arr.push(tmp);
        }
      }
    }
  }
  const rows = [...byRow.keys()].sort((a, b) => a - b);
  const rowXml: string[] = [];
  const defaultHtStr = rowHeightPxToOoxmlHt(sheet.defaultRowHeight);
  const defaultCwStr = colWidthPxToOoxmlWidth(sheet.defaultColWidth);
  const colsXml = buildColsXml(sheet);

  for (const r of rows) {
    const cells = byRow.get(r);
    if (cells === undefined) {
      continue;
    }
    cells.sort((a, b) => a.col - b.col);
    const spans =
      cells.length > 0 ? `${cells[0].col + 1}:${cells[cells.length - 1].col + 1}` : "1:1";
    const cXml = cells.map((c) => cellToXml(sheet, c, sst, xfBySig, opts)).join("");
    const rowHtAttr = rowXmlAttributesForHeight(sheet, r);
    rowXml.push(`<row r="${r + 1}" spans="${spans}"${rowHtAttr}>${cXml}</row>`);
  }

  const mergeXml = mergeCellsXml(sheet);

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<worksheet xmlns="${SS_MAIN}" xmlns:r="${REL_NS}">` +
    `<dimension ref="${dim}"/>` +
    `<sheetViews><sheetView workbookViewId="0" tabSelected="${sheetIndex === 0 ? "1" : "0"}"/></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="${defaultHtStr}" defaultColWidth="${defaultCwStr}"/>` +
    colsXml +
    `<sheetData>${rowXml.join("")}</sheetData>` +
    mergeXml +
    `</worksheet>`
  );
}

function sanitizeSheetName(name: string, index: number): string {
  const cleaned = name
    .replace(/[[\]*?:/\\]/g, "_")
    .trim()
    .slice(0, 31);
  if (cleaned === "") {
    return `Sheet${index + 1}`;
  }
  return cleaned;
}

/** 导出为标准 XLSX（ECMA-376 OPC + ZIP + deflate）。 */
export function exportWorkbookToXlsxBytes(
  workbook: Workbook,
  options: XlsxExportOptions = DEFAULT_XLSX_EXPORT_OPTIONS,
): Uint8Array {
  if (workbook.sheetCount === 0) {
    throw new Error("工作簿至少需一张工作表");
  }

  const styleTable = buildStyleTable(workbook, options);
  const stylesXml = buildStylesXml(styleTable);
  const { xml: sstXml, index: sstMap } = buildSharedStrings(workbook, options);

  const sheetParts: string[] = [];
  for (let i = 0; i < workbook.sheetCount; i++) {
    const sh = workbook.getSheet(i);
    if (sh === undefined) {
      continue;
    }
    sheetParts.push(buildSheetXml(sh, i, sstMap, styleTable.xfBySig, options));
  }

  const sheetNames = sheetParts.map((_, i) => {
    const sh = workbook.getSheet(i);
    return sanitizeSheetName(sh?.name ?? `Sheet${i + 1}`, i);
  });

  const n = sheetParts.length;
  const sheetRels: string[] = [];
  let rid = 1;
  for (let i = 0; i < n; i++) {
    sheetRels.push(
      `<Relationship Id="rId${rid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`,
    );
    rid++;
  }
  const ssRid = rid;
  sheetRels.push(
    `<Relationship Id="rId${ssRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>`,
  );
  rid++;
  const stRid = rid;
  sheetRels.push(
    `<Relationship Id="rId${stRid}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>`,
  );

  const sheetsXml = sheetNames
    .map((name, i) => `<sheet name="${escapeXml(name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`)
    .join("");

  const workbookXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<workbook xmlns="${SS_MAIN}" xmlns:r="${REL_NS}">` +
    `<workbookPr date1904="false"/>` +
    `<bookViews><workbookView xWindow="0" yWindow="0"/></bookViews>` +
    `<sheets>${sheetsXml}</sheets>` +
    `<calcPr calcId="191029"/>` +
    `</workbook>`;

  const workbookRelsXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="${PKG_REL}">${sheetRels.join("")}</Relationships>`;

  const rootRels =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="${PKG_REL}">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>` +
    `<Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/>` +
    `<Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/>` +
    `</Relationships>`;

  const ctOverrides: string[] = [
    `<Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/>`,
    `<Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/>`,
    `<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>`,
    `<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>`,
    `<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>`,
  ];
  for (let i = 0; i < n; i++) {
    ctOverrides.push(
      `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`,
    );
  }

  const contentTypes =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">` +
    `<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>` +
    `<Default Extension="xml" ContentType="application/xml"/>` +
    ctOverrides.join("") +
    `</Types>`;

  const coreXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<cp:coreProperties xmlns:cp="${CP_NS}" xmlns:dc="${DC_NS}" xmlns:dcterms="${DCTERMS_NS}" xmlns:xsi="${XSI_NS}">` +
    `<dc:creator>FlexSheet</dc:creator>` +
    `<cp:lastModifiedBy>FlexSheet</cp:lastModifiedBy>` +
    `<dcterms:created xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:created>` +
    `<dcterms:modified xsi:type="dcterms:W3CDTF">${new Date().toISOString()}</dcterms:modified>` +
    `</cp:coreProperties>`;

  const appXml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties" xmlns:vt="http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes">` +
    `<Application>Microsoft Excel</Application>` +
    `<DocSecurity>0</DocSecurity>` +
    `<ScaleCrop>false</ScaleCrop>` +
    `<HeadingPairs><vt:vector size="2" baseType="variant"><vt:variant><vt:lpstr>Worksheets</vt:lpstr></vt:variant><vt:variant><vt:i4>${n}</vt:i4></vt:variant></vt:vector></HeadingPairs>` +
    `<TitlesOfParts><vt:vector size="${n}" baseType="lpstr">${sheetNames.map((nm) => `<vt:lpstr>${escapeXml(nm)}</vt:lpstr>`).join("")}</vt:vector></TitlesOfParts>` +
    `</Properties>`;

  const enc = new TextEncoder();
  const files: ZipEntryInput[] = [
    { path: "[Content_Types].xml", data: enc.encode(contentTypes) },
    { path: "_rels/.rels", data: enc.encode(rootRels) },
    { path: "docProps/core.xml", data: enc.encode(coreXml) },
    { path: "docProps/app.xml", data: enc.encode(appXml) },
    { path: "xl/workbook.xml", data: enc.encode(workbookXml) },
    { path: "xl/_rels/workbook.xml.rels", data: enc.encode(workbookRelsXml) },
    { path: "xl/styles.xml", data: enc.encode(stylesXml) },
    { path: "xl/sharedStrings.xml", data: enc.encode(sstXml) },
  ];
  for (let i = 0; i < n; i++) {
    files.push({
      path: `xl/worksheets/sheet${i + 1}.xml`,
      data: enc.encode(sheetParts[i] ?? ""),
    });
  }

  return buildZipArchive(files);
}

export function exportWorkbookToXlsxBlob(workbook: Workbook, options?: XlsxExportOptions): Blob {
  const bytes = exportWorkbookToXlsxBytes(workbook, options ?? DEFAULT_XLSX_EXPORT_OPTIONS);
  return new Blob([new Uint8Array(bytes)], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}
