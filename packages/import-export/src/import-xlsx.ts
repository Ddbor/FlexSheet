import { Workbook } from "@flexsheet/core";
import { Worksheet } from "@flexsheet/core";
import type {
  CellBorderKind,
  CellBorderSide,
  CellScalar,
  CellStyle,
  CellTextOrientation,
} from "@flexsheet/core";
import { parseCellRef } from "./a1.js";
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
  const uEl = firstLocal(fontEl, "u");
  if (uEl !== undefined) {
    const uv = uEl.getAttribute("val");
    st.underline = uv === "double" ? "double" : "single";
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

function parseAlignmentFromXf(xf: Element): Partial<CellStyle> {
  const align = firstLocal(xf, "alignment");
  if (align === undefined) {
    return {};
  }
  const out: Partial<CellStyle> = {};
  const h = align.getAttribute("horizontal");
  if (h === "left" || h === "center" || h === "right") {
    out.hAlign = h;
  }
  const v = align.getAttribute("vertical");
  if (v === "top") {
    out.vAlign = "top";
  } else if (v === "center") {
    out.vAlign = "middle";
  } else if (v === "bottom") {
    out.vAlign = "bottom";
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
  const tr = align.getAttribute("textRotation");
  if (tr !== null && tr !== "") {
    const n = Number(tr);
    if (Number.isFinite(n)) {
      const o = parseTextOrientationFromOoxml(n);
      if (o !== undefined) {
        out.textOrientation = o;
      }
    }
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
      rgb !== null && rgb !== undefined && rgb !== ""
        ? { kind, colorArgb: rgb }
        : { kind };
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
  if (pf.getAttribute("patternType") !== "solid") {
    return {};
  }
  const fg = firstLocal(pf, "fgColor");
  const rgb = fg?.getAttribute("rgb");
  if (rgb !== null && rgb !== undefined && rgb !== "") {
    return { fillArgb: rgb };
  }
  return {};
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
  const borderEls =
    bordersEl !== undefined ? childrenLocal(bordersEl, "border") : [];
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

/** 自标准 XLSX Blob 导入为 `Workbook`（纯前端）。 */
export async function importXlsxToWorkbook(blob: Blob): Promise<Workbook> {
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
  }

  if (out.sheetCount === 0) {
    throw new Error("XLSX：未解析到任何工作表");
  }

  for (let i = 0; i < out.sheetCount; i++) {
    const sh = out.getSheet(i);
    if (sh !== undefined) {
      recalcWorksheet(sh);
    }
  }

  return out;
}
