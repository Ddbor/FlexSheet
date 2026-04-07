import type { Cell, CellScalar, CellStyle } from "@flexsheet/core";
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

function styleSignature(st: CellStyle | null | undefined): string {
  if (st === null || st === undefined) {
    return "";
  }
  return JSON.stringify({
    b: st.bold === true,
    fg: st.fgArgb ?? "",
    fill: st.fillArgb ?? "",
  });
}

interface StyleTable {
  readonly xfBySig: Map<string, number>;
  readonly fontsXml: string[];
  readonly fillsXml: string[];
  readonly cellXfsXml: string[];
}

/** `styleSignature` 与 `ensureStyle` 中 JSON 往返用的稳定形状。 */
interface StyleSignaturePayload {
  readonly b: boolean;
  readonly fg: string;
  readonly fill: string;
}

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
  const cellXfsXml: string[] = [`<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`];
  return { xfBySig, fontsXml, fillsXml, cellXfsXml };
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
  const cellXfsXml: string[] = [`<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>`];

  let nextFont = 1;
  let nextFill = 2;
  let nextXf = 1;

  const ensureStyle = (sig: string): void => {
    if (sig === "" || xfBySig.has(sig)) {
      return;
    }
    const st = JSON.parse(sig) as StyleSignaturePayload;
    let fontId = 0;
    const needFont = st.b || st.fg !== "";
    if (needFont) {
      const bold = st.b ? "<b/>" : "";
      const color = st.fg !== "" ? `<color rgb="${escapeXml(st.fg)}"/>` : `<color rgb="FF000000"/>`;
      fontsXml.push(
        `<font>${bold}<sz val="11"/>${color}<name val="Calibri"/><family val="2"/></font>`,
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

    const applyFont = needFont ? ` applyFont="1"` : "";
    const applyFill = st.fill !== "" ? ` applyFill="1"` : "";
    cellXfsXml.push(
      `<xf numFmtId="0" fontId="${fontId}" fillId="${fillId}" borderId="0" xfId="0"${applyFont}${applyFill}/>`,
    );
    xfBySig.set(sig, nextXf++);
  };

  for (let i = 0; i < workbook.sheetCount; i++) {
    const sh = workbook.getSheet(i);
    if (sh === undefined) {
      continue;
    }
    sh.iterateCells((c) => {
      if (!shouldExportCellForXlsx(c, opts)) {
        return;
      }
      ensureStyle(styleSignature(c.style));
    });
  }

  return { xfBySig, fontsXml, fillsXml, cellXfsXml };
}

function buildStylesXml(table: StyleTable): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<styleSheet xmlns="${SS_MAIN}">` +
    `<numFmts count="0"/>` +
    `<fonts count="${table.fontsXml.length}">${table.fontsXml.join("")}</fonts>` +
    `<fills count="${table.fillsXml.length}">${table.fillsXml.join("")}</fills>` +
    `<borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders>` +
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

function shouldExportCellForXlsx(cell: Cell, opts: XlsxExportOptions): boolean {
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

function usedBoundsFiltered(
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
    if (!shouldExportCellForXlsx(c, opts)) {
      return;
    }
    hasCell = true;
    minR = Math.min(minR, c.row);
    maxR = Math.max(maxR, c.row);
    minC = Math.min(minC, c.col);
    maxC = Math.max(maxC, c.col);
  });
  if (!hasCell) {
    return null;
  }
  return { minR, maxR, minC, maxC };
}

function rowHtPoints(sheet: Worksheet): string {
  const px = sheet.defaultRowHeight;
  const pt = (px * 72) / 96;
  return pt.toFixed(2);
}

function colWidthChars(sheet: Worksheet): string {
  const w = Math.max(0, (sheet.defaultColWidth - 5) / 7);
  return w.toFixed(2);
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
  cell: Cell,
  sst: Map<string, number>,
  xfBySig: Map<string, number>,
  opts: XlsxExportOptions,
): string {
  const ref = formatCellRef(cell.row, cell.col);
  const sig = opts.includeStyles ? styleSignature(cell.style) : "";
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
      if (!shouldExportCellForXlsx(c, opts)) {
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

function buildSheetXml(
  sheet: Worksheet,
  sheetIndex: number,
  sst: Map<string, number>,
  xfBySig: Map<string, number>,
  opts: XlsxExportOptions,
): string {
  const b = usedBoundsFiltered(sheet, opts);
  const dim =
    b === null ? "A1" : `${formatCellRef(b.minR, b.minC)}:${formatCellRef(b.maxR, b.maxC)}`;

  const byRow = new Map<number, Cell[]>();
  sheet.iterateCells((c) => {
    if (!shouldExportCellForXlsx(c, opts)) {
      return;
    }
    const arr = byRow.get(c.row);
    if (arr === undefined) {
      byRow.set(c.row, [c]);
    } else {
      arr.push(c);
    }
  });
  const rows = [...byRow.keys()].sort((a, b) => a - b);
  const rowXml: string[] = [];
  const ht = rowHtPoints(sheet);
  const cw = colWidthChars(sheet);
  const maxCol = b === null ? 1 : b.maxC + 1;

  for (const r of rows) {
    const cells = byRow.get(r);
    if (cells === undefined) {
      continue;
    }
    cells.sort((a, b) => a.col - b.col);
    const spans =
      cells.length > 0 ? `${cells[0].col + 1}:${cells[cells.length - 1].col + 1}` : "1:1";
    const cXml = cells.map((c) => cellToXml(c, sst, xfBySig, opts)).join("");
    rowXml.push(`<row r="${r + 1}" spans="${spans}" ht="${ht}" customHeight="1">${cXml}</row>`);
  }

  const colsXml =
    `<cols>` + `<col min="1" max="${maxCol}" width="${cw}" customWidth="1"/>` + `</cols>`;

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<worksheet xmlns="${SS_MAIN}" xmlns:r="${REL_NS}">` +
    `<dimension ref="${dim}"/>` +
    `<sheetViews><sheetView workbookViewId="0" tabSelected="${sheetIndex === 0 ? "1" : "0"}"/></sheetViews>` +
    `<sheetFormatPr defaultRowHeight="${ht}" defaultColWidth="${cw}"/>` +
    colsXml +
    `<sheetData>${rowXml.join("")}</sheetData>` +
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

export function exportWorkbookToXlsxBlob(
  workbook: Workbook,
  options?: XlsxExportOptions,
): Blob {
  const bytes = exportWorkbookToXlsxBytes(workbook, options ?? DEFAULT_XLSX_EXPORT_OPTIONS);
  return new Blob([new Uint8Array(bytes)], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
  });
}
