import { Workbook } from "@flexsheet/core";
import { Worksheet } from "@flexsheet/core";
import type { CellScalar, CellStyle } from "@flexsheet/core";
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
  const color = firstLocal(fontEl, "color");
  const rgb = color?.getAttribute("rgb");
  if (rgb !== null && rgb !== undefined && rgb !== "") {
    st.fgArgb = rgb;
  }
  return st;
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

function parseStylesTable(stylesXml: string | undefined): (CellStyle | null)[] {
  if (stylesXml === undefined) {
    return [];
  }
  const doc = parseXml(stylesXml);
  const root = doc.documentElement;
  const fontsEl = firstLocal(root, "fonts");
  const fillsEl = firstLocal(root, "fills");
  const cellXfsEl = firstLocal(root, "cellXfs");
  if (fontsEl === undefined || fillsEl === undefined || cellXfsEl === undefined) {
    return [];
  }
  const fontEls = childrenLocal(fontsEl, "font");
  const fillEls = childrenLocal(fillsEl, "fill");
  const fontParts = fontEls.map((f) => parseFontStyle(f));
  const fillParts = fillEls.map((f) => parseFillStyle(f));

  const xfs = childrenLocal(cellXfsEl, "xf");
  const table: (CellStyle | null)[] = [];
  for (const xf of xfs) {
    const fontId = Number(xf.getAttribute("fontId") ?? "0");
    const fillId = Number(xf.getAttribute("fillId") ?? "0");
    const applyFont = xf.getAttribute("applyFont") === "1";
    const applyFill = xf.getAttribute("applyFill") === "1";
    const st: CellStyle = {};
    if (applyFont && fontParts[fontId] !== undefined) {
      Object.assign(st, fontParts[fontId]);
    }
    if (applyFill && fillParts[fillId] !== undefined) {
      Object.assign(st, fillParts[fillId]);
    }
    Object.assign(st, parseAlignmentFromXf(xf));
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

  const sfmt = firstLocal(root, "sheetFormatPr");
  const defaultRowHt =
    sfmt !== undefined ? Number(sfmt.getAttribute("defaultRowHeight") ?? "15") : 15;
  const defaultColW = sfmt !== undefined ? Number(sfmt.getAttribute("defaultColWidth") ?? "9") : 9;

  const sheet = new Worksheet(name, 1000, 64);
  sheet.defaultRowHeight = (defaultRowHt * 96) / 72;
  sheet.defaultColWidth = defaultColW * 7 + 5;

  const data = firstLocal(root, "sheetData");
  if (data === undefined) {
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
    const nextR = Math.max(sheet.rowCount, maxR + 1);
    const nextC = Math.max(sheet.colCount, maxC + 1);
    sheet.rowCount = Math.max(1, nextR);
    sheet.colCount = Math.max(1, nextC);
  });
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
