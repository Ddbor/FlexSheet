import { Workbook } from "@flexsheet/core";
import { Worksheet } from "@flexsheet/core";
import type {
  CellBorderKind,
  CellBorderSide,
  CellScalar,
  CellStyle,
  CellTextOrientation,
  ParsedTableStyleCommand,
  PivotAggregateKind,
  PivotValueFieldSpec,
  WorksheetPivotTableDefinition,
} from "@flexsheet/core";
import { ooxmlTableStyleNameToParsed } from "@flexsheet/core";
import { isCellFillPatternType } from "@flexsheet/core";
import {
  isUnconfiguredPivotDefinition,
  writeUnconfiguredPivotPlaceholderToSheet,
} from "@flexsheet/core";
import { parseCellRef } from "./a1.js";
import {
  collectSheetFloatingPicturesFromXlsx,
  parseWorkbookThemeSchemeColors,
} from "./import-xlsx-drawing.js";
import type { XlsxImportedFloatingPicture } from "./import-xlsx-drawing.js";
import { unzipToMap } from "./zip-reader.js";
import { recalcWorksheet } from "@flexsheet/formula";

const REL_NS = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

function parseXml(xml: string): Document {
  const doc = new DOMParser().parseFromString(xml, "application/xml");
  const root = doc.documentElement;
  if (root === null || root.localName === "parsererror") {
    throw new Error("XLSX：XML 解析失败");
  }
  return doc;
}

function childrenLocal(el: Element, local: string): Element[] {
  return [...el.children].filter((c) => c.localName === local);
}

function firstLocal(el: Element, local: string): Element | undefined {
  return childrenLocal(el, local)[0];
}

function rId(el: Element): string | null {
  return el.getAttributeNS(REL_NS, "id") ?? el.getAttribute("r:id") ?? null;
}

function readSharedStrings(xml: string | undefined): string[] {
  if (xml === undefined) {
    return [];
  }
  const doc = parseXml(xml);
  const root = doc.documentElement;
  const out: string[] = [];
  for (const si of childrenLocal(root, "si")) {
    out.push(si.textContent ?? "");
  }
  return out;
}

function parseFontStyle(fontEl: Element): Partial<CellStyle> {
  const st: Partial<CellStyle> = {};
  if (childrenLocal(fontEl, "b").length > 0) {
    st.bold = true;
  }
  if (childrenLocal(fontEl, "i").length > 0) {
    st.italic = true;
  }
  if (childrenLocal(fontEl, "strike").length > 0) {
    st.strikethrough = true;
  }
  const uEl = firstLocal(fontEl, "u");
  if (uEl !== undefined) {
    const uv = uEl.getAttribute("val");
    st.underline = uv === "double" ? "double" : "single";
  }
  const vaEl = firstLocal(fontEl, "vertAlign");
  if (vaEl !== undefined) {
    const vv = vaEl.getAttribute("val");
    if (vv === "superscript" || vv === "subscript") {
      st.fontScript = vv;
    }
  }
  const szEl = firstLocal(fontEl, "sz");
  if (szEl !== undefined) {
    const v = Number(szEl.getAttribute("val"));
    if (Number.isFinite(v) && v > 0) {
      st.fontSizePt = v;
    }
  }
  const nameEl = firstLocal(fontEl, "name");
  if (nameEl !== undefined) {
    const nm = nameEl.getAttribute("val");
    if (nm !== null && nm !== "") {
      st.fontFamily = nm;
    }
  }
  const color = firstLocal(fontEl, "color");
  const rgb = color?.getAttribute("rgb");
  if (rgb !== null && rgb !== undefined && rgb !== "") {
    st.fgArgb = rgb;
  }
  return st;
}

function parseTextOrientationFromOoxml(n: number): CellTextOrientation | undefined {
  if (n === 45) {
    return "angleUp45";
  }
  if (n === 135) {
    return "angleDown45";
  }
  if (n === 255) {
    return "verticalStack";
  }
  if (n === 90) {
    return "rotateUp90";
  }
  if (n === 180) {
    return "rotateDown90";
  }
  return undefined;
}

/** OOXML `textRotation` → 画布用角度（°），逆时针为正。 */
function ooxmlTextRotationToDegrees(n: number): number | undefined {
  if (!Number.isFinite(n)) {
    return undefined;
  }
  if (n === 0) {
    return 0;
  }
  if (n > 0 && n <= 90) {
    return n;
  }
  if (n > 90 && n < 180) {
    return -(n - 90);
  }
  if (n === 180) {
    return -90;
  }
  return undefined;
}

function parseAlignmentFromXf(xf: Element): Partial<CellStyle> {
  const align = firstLocal(xf, "alignment");
  if (align === undefined) {
    return {};
  }
  const out: Partial<CellStyle> = {};
  const h = align.getAttribute("horizontal");
  if (
    h === "left" ||
    h === "center" ||
    h === "right" ||
    h === "fill" ||
    h === "justify" ||
    h === "distributed" ||
    h === "centerContinuous"
  ) {
    out.hAlign = h;
  }
  const v = align.getAttribute("vertical");
  if (v === "top") {
    out.vAlign = "top";
  } else if (v === "center") {
    out.vAlign = "middle";
  } else if (v === "bottom") {
    out.vAlign = "bottom";
  } else if (v === "justify") {
    out.vAlign = "justify";
  } else if (v === "distributed") {
    out.vAlign = "distributed";
  }
  const ind = align.getAttribute("indent");
  if (ind !== null && ind !== "") {
    const n = Number(ind);
    if (Number.isFinite(n) && n >= 0 && n <= 255) {
      const rounded = Math.round(n);
      if (rounded > 0) {
        out.indentLevel = rounded;
      }
    }
  }
  if (align.getAttribute("wrapText") === "1") {
    out.wrapText = true;
  }
  if (align.getAttribute("shrinkToFit") === "1") {
    out.shrinkToFit = true;
  }
  const tr = align.getAttribute("textRotation");
  if (tr !== null && tr !== "") {
    const n = Number(tr);
    if (Number.isFinite(n)) {
      if (n === 255) {
        out.textOrientation = "verticalStack";
      } else {
        const o = parseTextOrientationFromOoxml(n);
        if (o !== undefined) {
          out.textOrientation = o;
        } else {
          const deg = ooxmlTextRotationToDegrees(n);
          if (deg !== undefined && deg !== 0) {
            out.textRotationDegrees = deg;
          }
        }
      }
    }
  }
  return out;
}

function parseProtectionFromXf(xf: Element): Partial<CellStyle> {
  const prot = firstLocal(xf, "protection");
  if (prot === undefined) {
    return {};
  }
  const out: Partial<CellStyle> = {};
  const locked = prot.getAttribute("locked");
  if (locked === "0" || locked === "false") {
    out.locked = false;
  } else if (locked === "1" || locked === "true") {
    out.locked = true;
  }
  const hidden = prot.getAttribute("hidden");
  if (hidden === "1" || hidden === "true") {
    out.formulaHidden = true;
  } else if (hidden === "0" || hidden === "false") {
    out.formulaHidden = false;
  }
  return out;
}

const OOXML_BORDER_TO_KIND: Record<string, CellBorderKind> = {
  thin: "thin",
  medium: "medium",
  thick: "thick",
  double: "double",
  hairline: "hairline",
};

function ooxmlBorderKind(style: string): CellBorderKind {
  return OOXML_BORDER_TO_KIND[style] ?? "thin";
}

function parseBorderSidesFromBorderEl(borderEl: Element): Partial<CellStyle> {
  const out: Partial<CellStyle> = {};
  const one = (
    tag: "left" | "right" | "top" | "bottom",
    prop: "borderLeft" | "borderRight" | "borderTop" | "borderBottom",
  ): void => {
    const el = firstLocal(borderEl, tag);
    if (el === undefined) {
      return;
    }
    const style = el.getAttribute("style");
    if (style === null || style === "") {
      return;
    }
    const kind = ooxmlBorderKind(style);
    const colorEl = firstLocal(el, "color");
    const rgb = colorEl?.getAttribute("rgb");
    const side: CellBorderSide =
      rgb !== null && rgb !== undefined && rgb !== "" ? { kind, colorArgb: rgb } : { kind };
    out[prop] = side;
  };
  one("left", "borderLeft");
  one("right", "borderRight");
  one("top", "borderTop");
  one("bottom", "borderBottom");
  return out;
}

function parseFillStyle(fillEl: Element): Partial<CellStyle> {
  const pf = firstLocal(fillEl, "patternFill");
  if (pf === undefined) {
    return {};
  }
  const pt = pf.getAttribute("patternType") ?? "none";
  const fgEl = firstLocal(pf, "fgColor");
  const bgEl = firstLocal(pf, "bgColor");
  const fgRgb = fgEl?.getAttribute("rgb");
  const bgRgb = bgEl?.getAttribute("rgb");

  if (pt === "solid") {
    if (fgRgb !== null && fgRgb !== undefined && fgRgb !== "") {
      return { fillArgb: fgRgb };
    }
    return {};
  }
  if (pt === "none" || pt === "") {
    return {};
  }
  if (!isCellFillPatternType(pt)) {
    return {};
  }
  const out: Partial<CellStyle> = {
    fillPatternType: pt,
  };
  if (bgRgb !== null && bgRgb !== undefined && bgRgb !== "") {
    out.fillArgb = bgRgb;
  }
  if (fgRgb !== null && fgRgb !== undefined && fgRgb !== "") {
    out.fillPatternFgArgb = fgRgb;
  }
  return out;
}

/** Excel 内置 numFmtId → formatCode（子集；与导出/常见文件一致）。 */
const EXCEL_BUILTIN_NUMFMT_ID_TO_CODE: Readonly<Record<number, string>> = {
  0: "General",
  1: "0",
  2: "0.00",
  3: "#,##0",
  4: "#,##0.00",
  5: '"$"#,##0_);("$"#,##0)',
  6: '"$"#,##0.00_);("$"#,##0.00)',
  7: '"$"#,##0.00_);[Red]("$"#,##0.00)',
  8: '"$"#,##0.00_);("$"#,##0.00)',
  9: "0%",
  10: "0.00%",
  11: "0.00E+00",
  12: "# ?/?",
  13: "# ??/??",
  14: "m/d/yyyy",
  15: "d-mmm-yy",
  16: "d-mmm",
  17: "mmm-yy",
  18: "h:mm AM/PM",
  19: "h:mm:ss AM/PM",
  20: "h:mm",
  21: "h:mm:ss",
  22: "m/d/yyyy h:mm",
  37: "#,##0_);(#,##0)",
  38: "#,##0_);[Red](#,##0)",
  39: "#,##0.00_);(#,##0.00)",
  40: "#,##0.00_);[Red](#,##0.00)",
  42: '_("$"* #,##0.00_);_("$"* (#,##0.00_);_("$"* "-"??_);_(@_)',
  45: "mm:ss",
  46: "[h]:mm:ss",
  47: "##0.0E+0",
  48: "##0.0E+0",
  49: "@",
};

function parseCustomNumFmts(root: Element): Map<number, string> {
  const out = new Map<number, string>();
  const numFmtsEl = firstLocal(root, "numFmts");
  if (numFmtsEl === undefined) {
    return out;
  }
  for (const nm of childrenLocal(numFmtsEl, "numFmt")) {
    const id = Number(nm.getAttribute("numFmtId") ?? "");
    const code = nm.getAttribute("formatCode");
    if (Number.isFinite(id) && code !== null && code !== "") {
      out.set(id, code);
    }
  }
  return out;
}

function numberFormatFromNumFmtId(
  id: number,
  custom: ReadonlyMap<number, string>,
): string | undefined {
  if (!Number.isFinite(id) || id === 0) {
    return undefined;
  }
  const c = custom.get(id);
  if (c !== undefined) {
    return c;
  }
  const b = EXCEL_BUILTIN_NUMFMT_ID_TO_CODE[id];
  if (b !== undefined && b !== "General") {
    return b;
  }
  return undefined;
}

function parseStylesTable(stylesXml: string | undefined): (CellStyle | null)[] {
  if (stylesXml === undefined) {
    return [];
  }
  const doc = parseXml(stylesXml);
  const root = doc.documentElement;
  const fontsEl = firstLocal(root, "fonts");
  const fillsEl = firstLocal(root, "fills");
  const bordersEl = firstLocal(root, "borders");
  const cellXfsEl = firstLocal(root, "cellXfs");
  if (fontsEl === undefined || fillsEl === undefined || cellXfsEl === undefined) {
    return [];
  }
  const fontEls = childrenLocal(fontsEl, "font");
  const fillEls = childrenLocal(fillsEl, "fill");
  const borderEls = bordersEl !== undefined ? childrenLocal(bordersEl, "border") : [];
  const fontParts = fontEls.map((f) => parseFontStyle(f));
  const fillParts = fillEls.map((f) => parseFillStyle(f));
  const borderParts = borderEls.map((b) => parseBorderSidesFromBorderEl(b));
  const customNumFmt = parseCustomNumFmts(root);

  const xfs = childrenLocal(cellXfsEl, "xf");
  const table: (CellStyle | null)[] = [];
  for (const xf of xfs) {
    const fontId = Number(xf.getAttribute("fontId") ?? "0");
    const fillId = Number(xf.getAttribute("fillId") ?? "0");
    const borderId = Number(xf.getAttribute("borderId") ?? "0");
    const applyFont = xf.getAttribute("applyFont") === "1";
    const applyFill = xf.getAttribute("applyFill") === "1";
    const applyBorder = xf.getAttribute("applyBorder") === "1";
    const applyNumber = xf.getAttribute("applyNumberFormat") === "1";
    const numFmtId = Number(xf.getAttribute("numFmtId") ?? "0");
    const st: CellStyle = {};
    if (applyFont && fontParts[fontId] !== undefined) {
      Object.assign(st, fontParts[fontId]);
    }
    if (applyFill && fillParts[fillId] !== undefined) {
      Object.assign(st, fillParts[fillId]);
    }
    if (applyBorder && borderParts[borderId] !== undefined) {
      Object.assign(st, borderParts[borderId]);
    }
    Object.assign(st, parseAlignmentFromXf(xf));
    if (xf.getAttribute("applyProtection") === "1" || firstLocal(xf, "protection") !== undefined) {
      Object.assign(st, parseProtectionFromXf(xf));
    }
    if (applyNumber || numFmtId > 0) {
      const nf = numberFormatFromNumFmtId(numFmtId, customNumFmt);
      if (nf !== undefined) {
        st.numberFormat = nf;
      }
    }
    table.push(Object.keys(st).length > 0 ? st : null);
  }
  return table;
}

function parseCellValue(
  cel: Element,
  sst: string[],
): { value: CellScalar; formula: string | null } {
  const t = cel.getAttribute("t");
  const vEl = firstLocal(cel, "v");
  const fEl = firstLocal(cel, "f");
  const vText = vEl?.textContent ?? "";

  let formula: string | null = null;
  if (fEl !== undefined) {
    const ft = fEl.textContent ?? "";
    formula = ft === "" ? null : `=${ft}`;
  }

  if (formula !== null) {
    if (t === "s") {
      const idx = Number(vText);
      const cached = sst[idx];
      return { value: cached ?? "", formula };
    }
    if (t === "b") {
      return { value: vText === "1", formula };
    }
    if (t === "str" || t === "inlineStr") {
      return { value: vText, formula };
    }
    if (vText === "") {
      return { value: null, formula };
    }
    const n = Number(vText);
    if (!Number.isNaN(n) && vText.trim() !== "") {
      return { value: n, formula };
    }
    return { value: vText, formula };
  }

  if (t === "s") {
    const idx = Number(vText);
    return { value: sst[idx] ?? "", formula: null };
  }
  if (t === "b") {
    return { value: vText === "1", formula: null };
  }
  if (t === "str") {
    return { value: vText, formula: null };
  }
  if (vText === "") {
    return { value: null, formula: null };
  }
  const n = Number(vText);
  if (!Number.isNaN(n) && vText.trim() !== "") {
    return { value: n, formula: null };
  }
  return { value: vText, formula: null };
}

function parseMergeRef(ref: string): {
  masterRow: number;
  masterCol: number;
  rowSpan: number;
  colSpan: number;
} | null {
  const parts = ref.split(":");
  if (parts.length !== 2) {
    return null;
  }
  const a = parseCellRef(parts[0] ?? "");
  const b = parseCellRef(parts[1] ?? "");
  if (a === null || b === null) {
    return null;
  }
  const minR = Math.min(a.row, b.row);
  const maxR = Math.max(a.row, b.row);
  const minC = Math.min(a.col, b.col);
  const maxC = Math.max(a.col, b.col);
  return {
    masterRow: minR,
    masterCol: minC,
    rowSpan: maxR - minR + 1,
    colSpan: maxC - minC + 1,
  };
}

function parseMergeRegionsFromSheet(root: Element): readonly {
  readonly masterRow: number;
  readonly masterCol: number;
  readonly rowSpan: number;
  readonly colSpan: number;
}[] {
  const mergeRoot = firstLocal(root, "mergeCells");
  if (mergeRoot === undefined) {
    return [];
  }
  const out: {
    masterRow: number;
    masterCol: number;
    rowSpan: number;
    colSpan: number;
  }[] = [];
  for (const mc of childrenLocal(mergeRoot, "mergeCell")) {
    const ref = mc.getAttribute("ref");
    if (ref === null) {
      continue;
    }
    const p = parseMergeRef(ref);
    if (p !== null && (p.rowSpan > 1 || p.colSpan > 1)) {
      out.push(p);
    }
  }
  return out;
}

function parseSheet(
  xml: string,
  sst: string[],
  styleTable: (CellStyle | null)[],
  name: string,
): Worksheet {
  const doc = parseXml(xml);
  const root = doc.documentElement;
  let maxR = 0;
  let maxC = 0;

  const mergeRegions = parseMergeRegionsFromSheet(root);

  const sfmt = firstLocal(root, "sheetFormatPr");
  const defaultRowHt =
    sfmt !== undefined ? Number(sfmt.getAttribute("defaultRowHeight") ?? "15") : 15;
  const defaultColW = sfmt !== undefined ? Number(sfmt.getAttribute("defaultColWidth") ?? "9") : 9;

  const sheet = new Worksheet(name, 1000, 64);
  sheet.defaultRowHeight = (defaultRowHt * 96) / 72;
  sheet.defaultColWidth = defaultColW * 7 + 5;

  const data = firstLocal(root, "sheetData");
  if (data === undefined) {
    sheet.restoreMergeRegionsFromSnapshot(mergeRegions);
    sheet.notifyDataChanged();
    return sheet;
  }

  sheet.batch(() => {
    for (const row of childrenLocal(data, "row")) {
      for (const cel of childrenLocal(row, "c")) {
        const ref = cel.getAttribute("r");
        if (ref === null) {
          continue;
        }
        const addr = parseCellRef(ref);
        if (addr === null) {
          continue;
        }
        const { row: rr, col: cc } = addr;
        maxR = Math.max(maxR, rr);
        maxC = Math.max(maxC, cc);
        const { value, formula } = parseCellValue(cel, sst);
        const cell = sheet.getCell(rr, cc);
        cell.formula = formula;
        cell.value = value;
        const sIdx = cel.getAttribute("s");
        if (sIdx !== null) {
          const ix = Number(sIdx);
          const st = styleTable[ix];
          if (st !== null && st !== undefined) {
            cell.style = { ...st };
          }
        }
      }
    }
    for (const r of mergeRegions) {
      maxR = Math.max(maxR, r.masterRow + r.rowSpan - 1);
      maxC = Math.max(maxC, r.masterCol + r.colSpan - 1);
    }
    const nextR = Math.max(sheet.rowCount, maxR + 1);
    const nextC = Math.max(sheet.colCount, maxC + 1);
    sheet.rowCount = Math.max(1, nextR);
    sheet.colCount = Math.max(1, nextC);

    for (const row of childrenLocal(data, "row")) {
      const rAttr = row.getAttribute("r");
      if (rAttr === null) {
        continue;
      }
      const rowNum = Number(rAttr);
      if (!Number.isFinite(rowNum) || rowNum < 1) {
        continue;
      }
      const rr = rowNum - 1;
      if (rr < 0 || rr >= sheet.rowCount) {
        continue;
      }
      const ht = row.getAttribute("ht");
      if (ht !== null && ht !== "") {
        const pt = Number(ht);
        if (Number.isFinite(pt) && pt > 0) {
          sheet.setRowHeight(rr, Math.max(2, Math.round((pt * 96) / 72)));
        }
      }
    }
  });
  sheet.batch(() => {
    const colsRoot = firstLocal(root, "cols");
    if (colsRoot === undefined) {
      return;
    }
    for (const col of childrenLocal(colsRoot, "col")) {
      const min = Number(col.getAttribute("min"));
      const max = Number(col.getAttribute("max"));
      const wAttr = col.getAttribute("width");
      if (!Number.isFinite(min) || !Number.isFinite(max) || wAttr === null || wAttr === "") {
        continue;
      }
      const charW = Number(wAttr);
      if (!Number.isFinite(charW) || charW <= 0) {
        continue;
      }
      const px = Math.max(2, Math.round(charW * 7 + 5));
      const minCol = Math.trunc(min) - 1;
      const maxCol = Math.trunc(max) - 1;
      for (let cc = minCol; cc <= maxCol; cc++) {
        if (cc >= 0 && cc < sheet.colCount) {
          sheet.setColWidth(cc, px);
        }
      }
    }
  });
  sheet.restoreMergeRegionsFromSnapshot(mergeRegions);
  sheet.notifyDataChanged();
  return sheet;
}

/** OOXML `workbookView/@activeTab`：与工作表列表顺序一致的从 0 开始的可见活动表索引。 */
function readActiveSheetIndexFromWorkbookXml(wbDoc: Document, sheetCount: number): number {
  if (sheetCount <= 0) {
    return 0;
  }
  const bookViews = firstLocal(wbDoc.documentElement, "bookViews");
  if (bookViews === undefined) {
    return 0;
  }
  const wbView = firstLocal(bookViews, "workbookView");
  if (wbView === undefined) {
    return 0;
  }
  const raw = wbView.getAttribute("activeTab");
  if (raw === null || raw.trim().length === 0) {
    return 0;
  }
  const n = Number.parseInt(raw, 10);
  if (!Number.isFinite(n) || n < 0) {
    return 0;
  }
  return Math.min(n, sheetCount - 1);
}

function resolveWorkbookRels(relsXml: string): Map<string, string> {
  const doc = parseXml(relsXml);
  const map = new Map<string, string>();
  for (const rel of childrenLocal(doc.documentElement, "Relationship")) {
    const id = rel.getAttribute("Id");
    const target = rel.getAttribute("Target");
    if (id !== null && target !== null) {
      map.set(id, target.replace(/^\//, ""));
    }
  }
  return map;
}

/** 从 .rels XML 中提取指定类型的所有 Target 路径。 */
function findRelTargetsByType(relsXml: string, type: string): string[] {
  const doc = parseXml(relsXml);
  const targets: string[] = [];
  for (const rel of childrenLocal(doc.documentElement, "Relationship")) {
    if (rel.getAttribute("Type") === type) {
      const target = rel.getAttribute("Target");
      if (target !== null) {
        targets.push(target);
      }
    }
  }
  return targets;
}

const REL_TABLE = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table";
const REL_PIVOT_TABLE_IMPORT =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";

/**
 * 扫描所有工作表的 .rels，建立 Table 名称 → { sheetName, ref } 映射，
 * 供 parsePivotCacheMeta 处理 worksheetSource name="..." 格式时使用。
 */
function buildTableNameToSheetRange(
  files: ReadonlyMap<string, Uint8Array>,
  sheetBindings: readonly { readonly sheetPath: string; readonly sheetName: string }[],
): Map<string, { sheetName: string; ref: string }> {
  const result = new Map<string, { sheetName: string; ref: string }>();
  for (const binding of sheetBindings) {
    const relsPath = relsPartPathFor(binding.sheetPath);
    const relsBytes = files.get(relsPath);
    if (relsBytes === undefined) continue;
    const relsXml = new TextDecoder("utf-8").decode(relsBytes);
    const tableTargets = findRelTargetsByType(relsXml, REL_TABLE);
    for (const target of tableTargets) {
      const tablePath = resolvePartTarget(binding.sheetPath, target);
      const tableBytes = files.get(tablePath);
      if (tableBytes === undefined) continue;
      const tableXml = new TextDecoder("utf-8").decode(tableBytes);
      const tableDoc = parseXml(tableXml);
      const tableEl = tableDoc.documentElement;
      const tableName = tableEl.getAttribute("name");
      const tableRef = tableEl.getAttribute("ref");
      if (tableName !== null && tableName !== "" && tableRef !== null && tableRef !== "") {
        result.set(tableName, { sheetName: binding.sheetName, ref: tableRef });
      }
    }
  }
  return result;
}

function normalizePartPath(path: string): string {
  const raw = path.replace(/\\/g, "/").replace(/^\/+/, "");
  const segs = raw.split("/");
  const out: string[] = [];
  for (const seg of segs) {
    if (seg === "" || seg === ".") {
      continue;
    }
    if (seg === "..") {
      if (out.length > 0) {
        out.pop();
      }
      continue;
    }
    out.push(seg);
  }
  return out.join("/");
}

function dirname(path: string): string {
  const p = normalizePartPath(path);
  const idx = p.lastIndexOf("/");
  if (idx < 0) {
    return "";
  }
  return p.slice(0, idx);
}

function resolvePartTarget(basePartPath: string, target: string): string {
  if (target.startsWith("/")) {
    return normalizePartPath(target.slice(1));
  }
  const baseDir = dirname(basePartPath);
  const joined = baseDir === "" ? target : `${baseDir}/${target}`;
  return normalizePartPath(joined);
}

function relsPartPathFor(partPath: string): string {
  const dir = dirname(partPath);
  const file = partPath.slice(partPath.lastIndexOf("/") + 1);
  return normalizePartPath(`${dir}/_rels/${file}.rels`);
}

function parseRangeA1(
  ref: string,
): { startRow: number; endRow: number; startCol: number; endCol: number } | null {
  const raw = ref.trim();
  if (raw.length === 0) {
    return null;
  }
  const parts = raw.split(":");
  const a = parseCellRef(parts[0] ?? "");
  const b = parseCellRef(parts[1] ?? parts[0] ?? "");
  if (a === null || b === null) {
    return null;
  }
  return {
    startRow: Math.min(a.row, b.row),
    endRow: Math.max(a.row, b.row),
    startCol: Math.min(a.col, b.col),
    endCol: Math.max(a.col, b.col),
  };
}

function parsePivotAggregateKind(subtotal: string | null): PivotAggregateKind {
  switch ((subtotal ?? "").trim().toLowerCase()) {
    case "count":
      return "count";
    case "average":
      return "average";
    case "max":
      return "max";
    case "min":
      return "min";
    case "sum":
    default:
      return "sum";
  }
}

function parsePivotFieldOffsets(container: Element | undefined): number[] {
  if (container === undefined) {
    return [];
  }
  const out: number[] = [];
  for (const f of [...container.children]) {
    if (f.localName === "field") {
      const x = Number(f.getAttribute("x") ?? "");
      if (Number.isFinite(x) && x >= 0) {
        out.push(Math.trunc(x));
      }
      continue;
    }
    if (f.localName === "pageField") {
      const fld = Number(f.getAttribute("fld") ?? "");
      if (Number.isFinite(fld) && fld >= 0) {
        out.push(Math.trunc(fld));
      }
    }
  }
  return out;
}

function inferPivotPageFilterStartRow(
  destSheet: Worksheet,
  pivotInnerTopRow: number,
  pivotInnerLeftCol: number,
  filterFieldCount: number,
): number | undefined {
  if (filterFieldCount <= 0 || pivotInnerTopRow <= 0) {
    return undefined;
  }
  const rows: number[] = [];
  for (let r = 0; r < pivotInnerTopRow; r++) {
    const a = destSheet.getCell(r, pivotInnerLeftCol).value;
    const b = destSheet.getCell(r, pivotInnerLeftCol + 1).value;
    const sa = a === null || a === "" ? "" : String(a).trim();
    const sb = b === null || b === "" ? "" : String(b).trim();
    if (sa !== "" || sb !== "") {
      rows.push(r);
    }
  }
  if (rows.length === 0) {
    return undefined;
  }
  return rows[0];
}

function pivotHeaderCaptionAt(sheet: Worksheet, headerRow: number, col: number): string {
  const v = sheet.getCell(headerRow, col).value;
  if (typeof v === "string") {
    return v.trim();
  }
  if (typeof v === "number" && Number.isFinite(v)) {
    return String(v);
  }
  return "";
}

function inferPivotDataNumberFormatFromHeaderCaption(caption: string): string | undefined {
  const t = caption.replace(/\s/g, "");
  if (t === "") {
    return undefined;
  }
  if (/花费占比|占比项/.test(t) || (/占比/.test(t) && !/点击/.test(t))) {
    return "0.00%";
  }
  if (/点击率/.test(t) || /CTR/i.test(t)) {
    return "0.00%";
  }
  if (/平均点击花费|CPC/i.test(t)) {
    return "0.0000";
  }
  if (/求和项:花费|花费/.test(t) && !/占比/.test(t)) {
    return "#,##0.00";
  }
  if (/点击量|展现量/.test(t)) {
    return "#,##0";
  }
  return undefined;
}

/** Excel 透视表缓存区样式：表头/行标签/总计行与列数字格式（与 PivotStyleLight16 接近）。 */
function applyPivotImportedPresentation(
  sheet: Worksheet,
  _def: WorksheetPivotTableDefinition,
  pivotTableXml: string,
): void {
  const doc = parseXml(pivotTableXml);
  const root = doc.documentElement;
  if (root.localName !== "pivotTableDefinition") {
    return;
  }
  const loc = firstLocal(root, "location");
  const ref = loc?.getAttribute("ref") ?? "";
  const dstRange = parseRangeA1(ref);
  if (dstRange === null) {
    return;
  }
  const firstHdr = Number(loc?.getAttribute("firstHeaderRow") ?? "0");
  if (!Number.isFinite(firstHdr) || firstHdr < 0) {
    return;
  }
  const headerRow = dstRange.startRow + Math.trunc(firstHdr);
  const innerLast = dstRange.endRow;
  const headerFill = "FFBDD7EE";
  const rowLabelFill = "FFF2F2F2";
  const totalFill = "FFD9E1F2";

  sheet.batch(() => {
    for (let r = dstRange.startRow; r <= dstRange.endRow; r++) {
      const isGrand = r === innerLast && dstRange.endRow > dstRange.startRow;
      const isHeader = r === headerRow;
      for (let c = dstRange.startCol; c <= dstRange.endCol; c++) {
        const cell = sheet.getCell(r, c);
        const prev =
          cell.style !== undefined && cell.style !== null ? { ...cell.style } : ({} as CellStyle);
        if (isHeader) {
          prev.bold = true;
          prev.fillArgb = headerFill;
          prev.hAlign = c === dstRange.startCol ? "left" : "center";
          prev.vAlign = "middle";
        } else if (isGrand) {
          prev.bold = true;
          prev.fillArgb = totalFill;
          prev.hAlign = c === dstRange.startCol ? "left" : "right";
          prev.vAlign = "middle";
        } else {
          if (c === dstRange.startCol) {
            prev.fillArgb = rowLabelFill;
          }
          prev.hAlign = c === dstRange.startCol ? "left" : "right";
          prev.vAlign = "middle";
        }
        if (!isHeader) {
          const cap = pivotHeaderCaptionAt(sheet, headerRow, c);
          const nf = inferPivotDataNumberFormatFromHeaderCaption(cap);
          if (nf !== undefined) {
            prev.numberFormat = nf;
          }
        }
        cell.style = prev;
      }
    }
  });
}

function applyImportedTablesFromSheetRels(
  files: ReadonlyMap<string, Uint8Array>,
  binding: { readonly sheetPath: string },
  sheet: Worksheet,
): void {
  const relsPath = relsPartPathFor(binding.sheetPath);
  const relsBytes = files.get(relsPath);
  if (relsBytes === undefined) {
    return;
  }
  const relsXml = new TextDecoder("utf-8").decode(relsBytes);
  const tableTargets = findRelTargetsByType(relsXml, REL_TABLE);
  for (const target of tableTargets) {
    const tablePath = resolvePartTarget(binding.sheetPath, target);
    const tableBytes = files.get(tablePath);
    if (tableBytes === undefined) {
      continue;
    }
    const tableXml = new TextDecoder("utf-8").decode(tableBytes);
    const tableDoc = parseXml(tableXml);
    const tableEl = tableDoc.documentElement;
    if (tableEl.localName !== "table") {
      continue;
    }
    const tableRef = tableEl.getAttribute("ref");
    if (tableRef === null || tableRef === "") {
      continue;
    }
    const rng = parseRangeA1(tableRef);
    if (rng === null) {
      continue;
    }
    const styleEl = firstLocal(tableEl, "tableStyleInfo");
    const styleName = styleEl?.getAttribute("name");
    let parsed: ParsedTableStyleCommand | null = ooxmlTableStyleNameToParsed(styleName);
    if (parsed === null) {
      parsed = { section: "light", row: 0, col: 1 };
    }
    const hasHeaders = true;
    sheet.registerTableStyleRegion(
      {
        startRow: rng.startRow,
        endRow: rng.endRow,
        startCol: rng.startCol,
        endCol: rng.endCol,
      },
      parsed,
      hasHeaders,
    );
    if (rng.endRow > rng.startRow) {
      sheet.applyImportedTableColumnAutoFilters({
        headerRow: rng.startRow,
        bodyRowStart: rng.startRow + 1,
        bodyRowEnd: rng.endRow,
        startCol: rng.startCol,
        endCol: rng.endCol,
      });
    }
  }
}

function parsePivotFilterSelectedKeys(
  pivotFieldEl: Element | undefined,
  sharedItems: readonly string[],
): readonly string[] {
  if (pivotFieldEl === undefined) {
    return [];
  }
  const items = firstLocal(pivotFieldEl, "items");
  if (items === undefined) {
    return [];
  }
  const explicit: string[] = [];
  let hasDefault = false;
  for (const it of childrenLocal(items, "item")) {
    if ((it.getAttribute("t") ?? "").toLowerCase() === "default") {
      hasDefault = true;
      continue;
    }
    if ((it.getAttribute("h") ?? "") === "1") {
      continue;
    }
    const x = Number(it.getAttribute("x") ?? "");
    if (!Number.isFinite(x) || x < 0) {
      continue;
    }
    const key = sharedItems[Math.trunc(x)];
    if (key !== undefined) {
      explicit.push(key);
    }
  }
  if (hasDefault || explicit.length === 0) {
    return [];
  }
  return explicit;
}

function pivotFieldNameFromCellValue(value: CellScalar, fallback: string): string {
  if (typeof value === "string") {
    const t = value.trim();
    if (t.length > 0) {
      return t;
    }
  }
  if (typeof value === "number" && Number.isFinite(value)) {
    return String(value);
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  return fallback;
}

interface PivotCacheMeta {
  readonly sourceSheetName: string;
  readonly sourceRange: { startRow: number; endRow: number; startCol: number; endCol: number };
  readonly cacheFieldNames: readonly string[];
  readonly sharedItemsByFieldOffset: ReadonlyMap<number, readonly string[]>;
}

function parsePivotCacheMeta(
  defXml: string,
  tableNameToInfo?: ReadonlyMap<string, { sheetName: string; ref: string }>,
): PivotCacheMeta | null {
  const doc = parseXml(defXml);
  const root = doc.documentElement;
  if (root.localName !== "pivotCacheDefinition") {
    return null;
  }
  const cacheSource = firstLocal(root, "cacheSource");
  const wsSource =
    cacheSource !== undefined ? firstLocal(cacheSource, "worksheetSource") : undefined;
  let sourceSheetName = wsSource?.getAttribute("sheet") ?? "";
  let sourceRef = wsSource?.getAttribute("ref") ?? "";

  // 处理 worksheetSource name="tableName" 格式（Excel 命名表格引用）
  if ((sourceSheetName === "" || sourceRef === "") && wsSource !== undefined) {
    const tableName = wsSource.getAttribute("name");
    if (tableName !== null && tableName !== "" && tableNameToInfo !== undefined) {
      const info = tableNameToInfo.get(tableName);
      if (info !== undefined) {
        sourceSheetName = info.sheetName;
        sourceRef = info.ref;
      }
    }
  }

  const sourceRange = parseRangeA1(sourceRef);
  if (sourceSheetName.trim() === "" || sourceRange === null) {
    return null;
  }

  const sharedItemsByFieldOffset = new Map<number, readonly string[]>();
  const cacheFieldNames: string[] = [];
  const cacheFields = firstLocal(root, "cacheFields");
  const fieldEls = cacheFields !== undefined ? childrenLocal(cacheFields, "cacheField") : [];
  for (let i = 0; i < fieldEls.length; i++) {
    const cf = fieldEls[i]!;
    cacheFieldNames.push(cf.getAttribute("name") ?? `Column${i + 1}`);
    const shared = firstLocal(cf, "sharedItems");
    const keys: string[] = [];
    if (shared !== undefined) {
      for (const k of [...shared.children]) {
        if (k.localName === "s") {
          keys.push(k.getAttribute("v") ?? "");
        } else if (k.localName === "n") {
          keys.push(k.getAttribute("v") ?? "");
        } else if (k.localName === "b") {
          keys.push((k.getAttribute("v") ?? "") === "1" ? "TRUE" : "FALSE");
        } else if (k.localName === "m") {
          keys.push("(空白)");
        }
      }
    }
    sharedItemsByFieldOffset.set(i, keys);
  }
  return { sourceSheetName, sourceRange, cacheFieldNames, sharedItemsByFieldOffset };
}

function parseWorksheetPivotDefinitions(
  pivotTableXml: string,
  cache: PivotCacheMeta,
  sourceSheetIndex: number,
  sourceSheet: Worksheet | undefined,
  destSheet: Worksheet,
  idSeed: string,
): WorksheetPivotTableDefinition[] {
  const doc = parseXml(pivotTableXml);
  const root = doc.documentElement;
  if (root.localName !== "pivotTableDefinition") {
    return [];
  }
  const location = firstLocal(root, "location");
  const ref = location?.getAttribute("ref") ?? "";
  const dstRange = parseRangeA1(ref);
  if (dstRange === null) {
    return [];
  }

  const sourceRange = cache.sourceRange;
  const toAbsCol = (offset: number): number => sourceRange.startCol + offset;
  const rowOffsets = parsePivotFieldOffsets(firstLocal(root, "rowFields"));
  const colOffsets = parsePivotFieldOffsets(firstLocal(root, "colFields"));
  const filterOffsets = parsePivotFieldOffsets(firstLocal(root, "pageFields"));
  const dataFields = childrenLocal(firstLocal(root, "dataFields") ?? root, "dataField");
  const valueFields: PivotValueFieldSpec[] = [];
  for (const df of dataFields) {
    const fld = Number(df.getAttribute("fld") ?? "");
    if (!Number.isFinite(fld) || fld < 0) {
      continue;
    }
    valueFields.push({
      col: toAbsCol(Math.trunc(fld)),
      aggregate: parsePivotAggregateKind(df.getAttribute("subtotal")),
    });
  }

  const pivotFieldsRoot = firstLocal(root, "pivotFields");
  const pivotFieldEls =
    pivotFieldsRoot !== undefined ? childrenLocal(pivotFieldsRoot, "pivotField") : [];
  const filterSelectedKeys: (readonly string[])[] = [];
  for (const off of filterOffsets) {
    const pField = pivotFieldEls[off];
    const sharedItems = cache.sharedItemsByFieldOffset.get(off) ?? [];
    filterSelectedKeys.push(parsePivotFilterSelectedKeys(pField, sharedItems));
  }

  let hasHeaders = true;
  if (sourceSheet !== undefined) {
    const fields = cache.cacheFieldNames;
    if (fields.length > 0) {
      let allMatch = true;
      for (let i = 0; i < fields.length; i++) {
        const absCol = sourceRange.startCol + i;
        if (absCol > sourceRange.endCol) {
          break;
        }
        const cellName = pivotFieldNameFromCellValue(
          sourceSheet.getCell(sourceRange.startRow, absCol).value,
          `Column${i + 1}`,
        );
        if (cellName !== fields[i]) {
          allMatch = false;
          break;
        }
      }
      hasHeaders = allMatch;
    }
  }

  const name = root.getAttribute("name") ?? "PivotTable";
  const valueFieldsOnRows =
    valueFields.length > 1 &&
    colOffsets.length === 0 &&
    (root.getAttribute("dataOnRows") ?? "") === "1";

  const pageFilterStart = inferPivotPageFilterStartRow(
    destSheet,
    dstRange.startRow,
    dstRange.startCol,
    filterOffsets.length,
  );
  const layoutStart =
    pageFilterStart !== undefined
      ? Math.min(pageFilterStart, dstRange.startRow)
      : dstRange.startRow;
  const outputRowCount = Math.max(1, dstRange.endRow - layoutStart + 1);
  const outputColCount = Math.max(1, dstRange.endCol - dstRange.startCol + 1);

  return [
    {
      id: `pvt-import-${idSeed}`,
      name,
      sourceSheetIndex,
      sourceRange: { ...sourceRange },
      hasHeaders,
      rowFieldCols: rowOffsets.map(toAbsCol),
      columnFieldCols: colOffsets.map(toAbsCol),
      filterFieldCols: filterOffsets.map(toAbsCol),
      filterSelectedKeys,
      valueFields,
      valueFieldsOnRows: valueFieldsOnRows ? true : undefined,
      destinationRow: dstRange.startRow,
      destinationCol: dstRange.startCol,
      outputRowCount,
      outputColCount,
      pageFilterStartRow: pageFilterStart,
    },
  ];
}

/** 标准 XLSX 导入结果：单元格数据 + 工作表级浮动绘图（图片等）。 */
export interface XlsxImportResult {
  readonly workbook: Workbook;
  readonly floatingPictures: readonly XlsxImportedFloatingPicture[];
}

/**
 * 自标准 XLSX Blob 导入工作簿与浮动图片（DrawingML / `xl/drawings`）。
 * 若只需 `Workbook`，可使用 `importXlsxToWorkbook`。
 */
export async function importXlsx(blob: Blob): Promise<XlsxImportResult> {
  const buf = await blob.arrayBuffer();
  const files = unzipToMap(buf);

  const wbPath = "xl/workbook.xml";
  const wbBytes = files.get(wbPath);
  if (wbBytes === undefined) {
    throw new Error("XLSX：缺少 xl/workbook.xml");
  }
  const wbText = new TextDecoder("utf-8").decode(wbBytes);
  const wbDoc = parseXml(wbText);

  const relsBytes = files.get("xl/_rels/workbook.xml.rels");
  if (relsBytes === undefined) {
    throw new Error("XLSX：缺少 xl/_rels/workbook.xml.rels");
  }
  const relsText = new TextDecoder("utf-8").decode(relsBytes);
  const relsMap = resolveWorkbookRels(relsText);

  const sheetsEl = firstLocal(wbDoc.documentElement, "sheets");
  if (sheetsEl === undefined) {
    throw new Error("XLSX：工作簿无 sheets");
  }

  const sstText = files.get("xl/sharedStrings.xml");
  const sst = readSharedStrings(
    sstText !== undefined ? new TextDecoder("utf-8").decode(sstText) : undefined,
  );

  const stylesText = files.get("xl/styles.xml");
  const styleTable = parseStylesTable(
    stylesText !== undefined ? new TextDecoder("utf-8").decode(stylesText) : undefined,
  );

  const out = new Workbook();
  const sheetElements = childrenLocal(sheetsEl, "sheet");

  const importedSheetBindings: {
    readonly importedIndex: number;
    readonly sheetName: string;
    readonly sheetPath: string;
  }[] = [];
  for (const se of sheetElements) {
    const nm = se.getAttribute("name") ?? "Sheet";
    const id = rId(se);
    if (id === null) {
      continue;
    }
    const target = relsMap.get(id);
    if (target === undefined) {
      continue;
    }
    const fullPath = target.startsWith("xl/") ? target : `xl/${target}`;
    const part = files.get(fullPath);
    if (part === undefined) {
      continue;
    }
    const sheetXml = new TextDecoder("utf-8").decode(part);
    const ws = parseSheet(sheetXml, sst, styleTable, nm);
    out.addSheet(ws);
    importedSheetBindings.push({
      importedIndex: out.sheetCount - 1,
      sheetName: nm,
      sheetPath: fullPath,
    });
  }

  if (out.sheetCount === 0) {
    throw new Error("XLSX：未解析到任何工作表");
  }

  out.activeSheetIndex = readActiveSheetIndexFromWorkbookXml(wbDoc, out.sheetCount);

  const tableNameToInfo = buildTableNameToSheetRange(files, importedSheetBindings);

  const cacheIdToMeta = new Map<number, PivotCacheMeta>();
  const pivotCaches = firstLocal(wbDoc.documentElement, "pivotCaches");
  if (pivotCaches !== undefined) {
    for (const pc of childrenLocal(pivotCaches, "pivotCache")) {
      const cacheId = Number(pc.getAttribute("cacheId") ?? "");
      const relId = rId(pc);
      if (!Number.isFinite(cacheId) || relId === null) {
        continue;
      }
      const target = relsMap.get(relId);
      if (target === undefined) {
        continue;
      }
      const cacheDefPath = resolvePartTarget("xl/workbook.xml", target);
      const cacheDefBytes = files.get(cacheDefPath);
      if (cacheDefBytes === undefined) {
        continue;
      }
      const cacheDefXml = new TextDecoder("utf-8").decode(cacheDefBytes);
      const meta = parsePivotCacheMeta(cacheDefXml, tableNameToInfo);
      if (meta !== null) {
        cacheIdToMeta.set(Math.trunc(cacheId), meta);
      }
    }
  }

  const sheetNameToImportedIndex = new Map<string, number>();
  for (const binding of importedSheetBindings) {
    sheetNameToImportedIndex.set(binding.sheetName, binding.importedIndex);
  }
  let importedPivotSerial = 1;
  for (const binding of importedSheetBindings) {
    const sheetRelsPath = relsPartPathFor(binding.sheetPath);
    const sheetRelsBytes = files.get(sheetRelsPath);
    if (sheetRelsBytes === undefined) {
      continue;
    }
    const sheetRelsXml = new TextDecoder("utf-8").decode(sheetRelsBytes);
    const pivotTargets = findRelTargetsByType(sheetRelsXml, REL_PIVOT_TABLE_IMPORT);
    if (pivotTargets.length === 0) {
      continue;
    }
    const destSheet = out.getSheet(binding.importedIndex);
    if (destSheet === undefined) {
      continue;
    }
    for (const target of pivotTargets) {
      const pivotTablePath = resolvePartTarget(binding.sheetPath, target);
      const pivotTableBytes = files.get(pivotTablePath);
      if (pivotTableBytes === undefined) {
        continue;
      }
      const pivotTableXml = new TextDecoder("utf-8").decode(pivotTableBytes);
      const pivotDoc = parseXml(pivotTableXml);
      const cacheId = Number(pivotDoc.documentElement.getAttribute("cacheId") ?? "");
      if (!Number.isFinite(cacheId)) {
        continue;
      }
      const cacheMeta = cacheIdToMeta.get(Math.trunc(cacheId));
      if (cacheMeta === undefined) {
        continue;
      }
      const sourceSheetIndex = sheetNameToImportedIndex.get(cacheMeta.sourceSheetName);
      if (sourceSheetIndex === undefined) {
        continue;
      }
      const sourceSheet = out.getSheet(sourceSheetIndex);
      const defs = parseWorksheetPivotDefinitions(
        pivotTableXml,
        cacheMeta,
        sourceSheetIndex,
        sourceSheet,
        destSheet,
        String(importedPivotSerial++),
      );
      for (const def of defs) {
        destSheet.registerPivotTableDefinition(def);
        if (isUnconfiguredPivotDefinition(def)) {
          writeUnconfiguredPivotPlaceholderToSheet(destSheet, def);
        } else {
          applyPivotImportedPresentation(destSheet, def, pivotTableXml);
        }
      }
    }
  }

  for (const binding of importedSheetBindings) {
    const sh = out.getSheet(binding.importedIndex);
    if (sh !== undefined) {
      applyImportedTablesFromSheetRels(files, binding, sh);
    }
  }

  const themeScheme = parseWorkbookThemeSchemeColors(files);
  const floatingPictures: XlsxImportedFloatingPicture[] = [];
  for (const binding of importedSheetBindings) {
    const sh = out.getSheet(binding.importedIndex);
    if (sh === undefined) {
      continue;
    }
    floatingPictures.push(
      ...collectSheetFloatingPicturesFromXlsx(
        files,
        binding.sheetPath,
        sh,
        binding.sheetName,
        binding.importedIndex,
        themeScheme,
      ),
    );
  }

  for (let i = 0; i < out.sheetCount; i++) {
    const sh = out.getSheet(i);
    if (sh !== undefined) {
      recalcWorksheet(sh);
    }
  }

  return { workbook: out, floatingPictures };
}

/** 自标准 XLSX Blob 导入为 `Workbook`（纯前端）；不含浮动图时请用 `importXlsx`。 */
export async function importXlsxToWorkbook(blob: Blob): Promise<Workbook> {
  return (await importXlsx(blob)).workbook;
}
