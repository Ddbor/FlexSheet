import {
  getPivotValueFieldCaption,
  normalizeSelectionRange,
  type CellScalar,
  type PivotAggregateKind,
  type Workbook,
  type Worksheet,
  type WorksheetPivotTableDefinition,
} from "@flexsheet/core";
import { formatCellRef } from "./a1.js";
import { escapeXml } from "./xml-escape.js";

const SS_MAIN = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const PKG_REL = "http://schemas.openxmlformats.org/package/2006/relationships";
const REL_PIVOT_CACHE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCache";
const REL_PIVOT_CACHE_RECORDS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
const REL_PIVOT_TABLE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";

export interface PivotExportPiece {
  readonly cacheIdx: number;
  readonly cacheId: number;
  readonly destSheetIndex: number;
  readonly pivotTableName: string;
  readonly workbookRelTarget: string;
  readonly pivotCacheDefinitionXml: string;
  /** 与 `pivotCacheDefinition/@recordCount` 条数一致；每条 `<r>` 对应源表一行。 */
  readonly pivotCacheRecordsXml: string;
  readonly pivotCacheDefinitionRelsXml: string;
  readonly pivotTableXml: string;
  readonly pivotTableRelsXml: string;
}

export interface PivotSheetRelsPlan {
  readonly sheetIndex: number;
  readonly relsXml: string;
  readonly pivotTablesFragment: string;
}

function fieldLabelFromCell(value: CellScalar, fallback: string): string {
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

function collectFieldNames(
  sheet: Worksheet,
  def: WorksheetPivotTableDefinition,
): readonly string[] {
  const n = normalizeSelectionRange(def.sourceRange);
  const names: string[] = [];
  for (let c = n.startCol; c <= n.endCol; c++) {
    const fallback = `Column${c - n.startCol + 1}`;
    const nm = def.hasHeaders
      ? fieldLabelFromCell(sheet.getCell(n.startRow, c).value, fallback)
      : fallback;
    names.push(nm);
  }
  return dedupePivotFieldNames(names);
}

/**
 * Excel 要求 cacheField/@name 在透视缓存内唯一；重复列标题需加后缀（与 Excel「名称2」风格一致）。
 */
function dedupePivotFieldNames(names: readonly string[]): string[] {
  const used = new Set<string>();
  const out: string[] = [];
  for (const raw of names) {
    let base = raw.trim();
    if (base === "") {
      base = "Column";
    }
    let candidate = raw.trim() === "" ? base : raw.trim();
    let suffix = 1;
    while (used.has(candidate)) {
      suffix++;
      candidate = `${base}${suffix}`;
    }
    used.add(candidate);
    out.push(candidate);
  }
  return out;
}

/**
 * 与 `pivotCacheRecords` 中该列单元格类型一致；Excel 会校验 sharedItems 与记录行的对应关系，
 * 仅有 `containsBlank` 而无具体子项时，配合 `pivotField/items@x` 易导致「文件损坏无法修复」。
 */
function buildSharedItemsXmlForField(
  src: Worksheet,
  def: WorksheetPivotTableDefinition,
  fieldCol: number,
): string {
  const n = normalizeSelectionRange(def.sourceRange);
  const rStart = def.hasHeaders ? n.startRow + 1 : n.startRow;
  const rEnd = n.endRow;
  const numbers = new Set<number>();
  const strings = new Set<string>();
  let hasFalse = false;
  let hasTrue = false;
  let hasBlank = false;
  for (let r = rStart; r <= rEnd; r++) {
    const v = src.getCell(r, fieldCol).value;
    if (v === null || v === "") {
      hasBlank = true;
      continue;
    }
    if (typeof v === "number") {
      if (Number.isFinite(v)) {
        numbers.add(v);
      } else {
        hasBlank = true;
      }
      continue;
    }
    if (typeof v === "boolean") {
      if (v) {
        hasTrue = true;
      } else {
        hasFalse = true;
      }
      continue;
    }
    strings.add(String(v));
  }
  const parts: string[] = [];
  const numsSorted = [...numbers].sort((a, b) => a - b);
  for (const nv of numsSorted) {
    parts.push(`<n v="${String(nv)}"/>`);
  }
  const strsSorted = [...strings].sort((a, b) => a.localeCompare(b, undefined, { sensitivity: "accent" }));
  for (const s of strsSorted) {
    parts.push(`<s v="${escapeXml(s)}"/>`);
  }
  if (hasFalse) {
    parts.push(`<b v="0"/>`);
  }
  if (hasTrue) {
    parts.push(`<b v="1"/>`);
  }
  if (hasBlank) {
    parts.push(`<m/>`);
  }
  const mixed =
    (numbers.size > 0 && strings.size > 0) ||
    (numbers.size > 0 && (hasFalse || hasTrue)) ||
    (strings.size > 0 && (hasFalse || hasTrue)) ||
    (hasFalse && hasTrue);
  const attrs: string[] = [];
  if (hasBlank) {
    attrs.push('containsBlank="1"');
  }
  if (mixed) {
    attrs.push('containsSemiMixedTypes="1"');
  }
  const cnt = parts.length;
  if (cnt > 0) {
    attrs.push(`count="${cnt}"`);
  }
  const attrStr = attrs.length > 0 ? ` ${attrs.join(" ")}` : "";
  return `<sharedItems${attrStr}>${parts.join("")}</sharedItems>`;
}

/**
 * 透视结果占位矩形（与 `pivotTableDefinition/location/@ref` 一致）。
 * 行/列数异常（≤0）时按 1 处理，避免出现 `maxR < minR` 导致 worksheet `dimension` 与 OOXML 非法。
 */
export function pivotOutputExtents(def: WorksheetPivotTableDefinition): {
  minR: number;
  maxR: number;
  minC: number;
  maxC: number;
} {
  const r0 = def.destinationRow;
  const c0 = def.destinationCol;
  const rows = Math.max(1, def.outputRowCount);
  const cols = Math.max(1, def.outputColCount);
  return {
    minR: r0,
    maxR: r0 + rows - 1,
    minC: c0,
    maxC: c0 + cols - 1,
  };
}

function ooxmlDataSubtotal(agg: PivotAggregateKind): string {
  switch (agg) {
    case "sum":
      return "sum";
    case "count":
      return "count";
    case "average":
      return "average";
    case "max":
      return "max";
    case "min":
      return "min";
    default:
      return "sum";
  }
}

/**
 * 按数据源区域逐行写入缓存记录，使 `recordCount` / `count` 与 `<r>` 条数一致。
 * 仅 `count="0"` 且无 `<r>` 时，部分 Excel 版本会报损坏或无法修复。
 */
function buildPivotCacheRecordsXml(
  src: Worksheet,
  def: WorksheetPivotTableDefinition,
): { readonly xml: string; readonly recordCount: number } {
  const n = normalizeSelectionRange(def.sourceRange);
  const rStart = def.hasHeaders ? n.startRow + 1 : n.startRow;
  const rEnd = n.endRow;
  const rowsXml: string[] = [];
  let recordCount = 0;
  for (let r = rStart; r <= rEnd; r++) {
    const cells: string[] = [];
    for (let c = n.startCol; c <= n.endCol; c++) {
      cells.push(pivotRecordFragmentFromScalar(src.getCell(r, c).value));
    }
    rowsXml.push(`<r>${cells.join("")}</r>`);
    recordCount++;
  }
  const inner = rowsXml.join("");
  const xml =
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<pivotCacheRecords xmlns="${SS_MAIN}" count="${recordCount}">${inner}</pivotCacheRecords>`;
  return { xml, recordCount };
}

function pivotRecordFragmentFromScalar(value: CellScalar): string {
  if (value === null || value === "") {
    return "<m/>";
  }
  if (typeof value === "number") {
    if (!Number.isFinite(value)) {
      return "<m/>";
    }
    return `<n v="${String(value)}"/>`;
  }
  if (typeof value === "boolean") {
    return `<b v="${value ? 1 : 0}"/>`;
  }
  return `<s v="${escapeXml(String(value))}"/>`;
}

function buildPivotCacheDefinitionRelsXml(cacheIdx: number): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="${PKG_REL}">` +
    `<Relationship Id="rId1" Type="${REL_PIVOT_CACHE_RECORDS}" Target="pivotCacheRecords${cacheIdx}.xml"/>` +
    `</Relationships>`
  );
}

function buildPivotCacheDefinitionXml(
  def: WorksheetPivotTableDefinition,
  src: Worksheet,
  sheetNames: readonly string[],
  fieldNames: readonly string[],
  recordCount: number,
): string {
  const n = normalizeSelectionRange(def.sourceRange);
  const sheetName = sheetNames[def.sourceSheetIndex] ?? `Sheet${def.sourceSheetIndex + 1}`;
  const ref = `${formatCellRef(n.startRow, n.startCol)}:${formatCellRef(n.endRow, n.endCol)}`;
  const fieldsXml = fieldNames
    .map((nm, fi) => {
      const col = n.startCol + fi;
      const shared = buildSharedItemsXmlForField(src, def, col);
      return `<cacheField name="${escapeXml(nm)}">${shared}</cacheField>`;
    })
    .join("");
  const saveData = recordCount > 0 ? 1 : 0;
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<pivotCacheDefinition xmlns="${SS_MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
    ` refreshOnLoad="1" recordCount="${recordCount}" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" saveData="${saveData}">` +
    `<cacheSource type="worksheet">` +
    `<worksheetSource ref="${escapeXml(ref)}" sheet="${escapeXml(sheetName)}"/>` +
    `</cacheSource>` +
    `<cacheFields count="${fieldNames.length}">${fieldsXml}</cacheFields>` +
    `</pivotCacheDefinition>`
  );
}

function buildPivotTableDefinitionXml(
  def: WorksheetPivotTableDefinition,
  fieldNames: readonly string[],
  cacheId: number,
): string {
  const n = normalizeSelectionRange(def.sourceRange);
  const toOff = (absCol: number): number => absCol - n.startCol;

  const rowOffs = def.rowFieldCols.map(toOff);
  const colOffs = def.columnFieldCols.map(toOff);
  const valueSpecs = def.valueFields;

  const ext = pivotOutputExtents(def);
  const locRef = `${formatCellRef(ext.minR, ext.minC)}:${formatCellRef(ext.maxR, ext.maxC)}`;

  const dataFieldIndices = new Set<number>();
  for (const v of valueSpecs) {
    dataFieldIndices.add(toOff(v.col));
    if (v.computed?.kind === "bucketRatio") {
      dataFieldIndices.add(toOff(v.computed.denominatorCol));
    }
  }
  const rowOffSet = new Set(rowOffs);
  const colOffSet = new Set(colOffs);
  const filterOffs = def.filterFieldCols.map(toOff);
  const filterOffSet = new Set(filterOffs);

  const pivotFieldsXml: string[] = [];
  for (let fi = 0; fi < fieldNames.length; fi++) {
    if (dataFieldIndices.has(fi)) {
      pivotFieldsXml.push(`<pivotField dataField="1" showAll="0"/>`);
    } else if (filterOffSet.has(fi)) {
      pivotFieldsXml.push(
        `<pivotField axis="axisPage" showAll="1"><items count="1"><item x="0"/></items></pivotField>`,
      );
    } else if (rowOffSet.has(fi)) {
      pivotFieldsXml.push(
        `<pivotField axis="axisRow" showAll="1"><items count="1"><item x="0"/></items></pivotField>`,
      );
    } else if (colOffSet.has(fi)) {
      pivotFieldsXml.push(
        `<pivotField axis="axisCol" showAll="1"><items count="1"><item x="0"/></items></pivotField>`,
      );
    } else {
      pivotFieldsXml.push(`<pivotField/>`);
    }
  }

  const rowFieldsXml =
    `<rowFields count="${rowOffs.length}">` +
    rowOffs.map((o) => `<field x="${o}"/>`).join("") +
    `</rowFields>`;
  const rowItemsXml = `<rowItems count="1"><i><x v="0"/></i></rowItems>`;

  let colFieldsXml = "";
  let colItemsXml = `<colItems count="1"><i/></colItems>`;
  if (colOffs.length > 0) {
    colFieldsXml =
      `<colFields count="${colOffs.length}">` +
      colOffs.map((o) => `<field x="${o}"/>`).join("") +
      `</colFields>`;
  }

  const dataCaption = "Values";
  const dataFieldParts: string[] = [];
  for (const vf of valueSpecs) {
    const valOff = toOff(vf.col);
    const valueFieldName = fieldNames[valOff] ?? "Values";
    const denName =
      vf.computed?.kind === "bucketRatio"
        ? (fieldNames[toOff(vf.computed.denominatorCol)] ?? "")
        : undefined;
    const dataName = escapeXml(getPivotValueFieldCaption(vf, valueFieldName, denName));
    const subAgg: PivotAggregateKind = vf.computed !== undefined ? "sum" : vf.aggregate;
    dataFieldParts.push(
      `<dataField name="${dataName}" fld="${valOff}" subtotal="${escapeXml(
        ooxmlDataSubtotal(subAgg),
      )}"/>`,
    );
  }
  const dataFieldsXml = `<dataFields count="${dataFieldParts.length}">${dataFieldParts.join("")}</dataFields>`;

  const pageFieldsXml =
    filterOffs.length === 0
      ? ""
      : `<pageFields count="${filterOffs.length}">${filterOffs
          .map((o) => `<field x="${o}"/>`)
          .join("")}</pageFields>`;

  const styleXml = `<pivotTableStyleInfo name="PivotStyleLight16" showRowHeaders="1" showColHeaders="1" showRowStripes="0" showColStripes="0"/>`;

  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<pivotTableDefinition xmlns="${SS_MAIN}"` +
    ` name="${escapeXml(def.name)}"` +
    ` cacheId="${cacheId}"` +
    ` dataCaption="${escapeXml(dataCaption)}"` +
    ` rowGrandTotals="1"` +
    ` colGrandTotals="1"` +
    ` updatedVersion="6"` +
    ` minRefreshableVersion="3"` +
    ` useAutoFormatting="0"` +
    ` createdVersion="6"` +
    ` compactData="1"` +
    ` applyNumberFormats="0"` +
    ` applyBorderFormats="0"` +
    ` applyFontFormats="0"` +
    ` applyPatternFormats="0"` +
    ` applyAlignmentFormats="0"` +
    ` applyWidthHeightFormats="1">` +
    `<location ref="${escapeXml(locRef)}" firstHeaderRow="1" firstDataRow="1" firstDataCol="1"/>` +
    `<pivotFields count="${fieldNames.length}">${pivotFieldsXml.join("")}</pivotFields>` +
    rowFieldsXml +
    rowItemsXml +
    colFieldsXml +
    colItemsXml +
    pageFieldsXml +
    dataFieldsXml +
    styleXml +
    `</pivotTableDefinition>`
  );
}

function buildPivotTableRelsXml(cacheIdx: number): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<Relationships xmlns="${PKG_REL}">` +
    `<Relationship Id="rId1" Type="${REL_PIVOT_CACHE}" Target="../pivotCache/pivotCacheDefinition${cacheIdx}.xml"/>` +
    `</Relationships>`
  );
}

/**
 * 收集工作簿内全部透视导出部件（顺序：按工作表索引、再按注册顺序）。
 */
export function collectPivotExportPieces(
  workbook: Workbook,
  sheetNames: readonly string[],
): readonly PivotExportPiece[] {
  const out: PivotExportPiece[] = [];
  let cacheIdx = 0;
  for (let si = 0; si < workbook.sheetCount; si++) {
    const sh = workbook.getSheet(si);
    if (sh === undefined) {
      continue;
    }
    for (const def of sh.getPivotTableDefinitionsSnapshot()) {
      const src = workbook.getSheet(def.sourceSheetIndex);
      if (src === undefined) {
        continue;
      }
      cacheIdx++;
      const cacheId = cacheIdx - 1;
      const fieldNames = collectFieldNames(src, def);
      if (fieldNames.length === 0) {
        continue;
      }
      const { xml: pivotCacheRecordsXml, recordCount } = buildPivotCacheRecordsXml(src, def);
      const pivotCacheDefinitionXml = buildPivotCacheDefinitionXml(
        def,
        src,
        sheetNames,
        fieldNames,
        recordCount,
      );
      const pivotCacheDefinitionRelsXml = buildPivotCacheDefinitionRelsXml(cacheIdx);
      const pivotTableXml = buildPivotTableDefinitionXml(def, fieldNames, cacheId);
      const pivotTableRelsXml = buildPivotTableRelsXml(cacheIdx);
      out.push({
        cacheIdx,
        cacheId,
        destSheetIndex: si,
        pivotTableName: def.name,
        workbookRelTarget: `pivotCache/pivotCacheDefinition${cacheIdx}.xml`,
        pivotCacheDefinitionXml,
        pivotCacheRecordsXml,
        pivotCacheDefinitionRelsXml,
        pivotTableXml,
        pivotTableRelsXml,
      });
    }
  }
  return out;
}

export function buildPivotSheetPlans(
  pieces: readonly PivotExportPiece[],
): readonly PivotSheetRelsPlan[] {
  const bySheet = new Map<number, PivotExportPiece[]>();
  for (const p of pieces) {
    const arr = bySheet.get(p.destSheetIndex);
    if (arr === undefined) {
      bySheet.set(p.destSheetIndex, [p]);
    } else {
      arr.push(p);
    }
  }
  const plans: PivotSheetRelsPlan[] = [];
  for (const [sheetIndex, list] of bySheet) {
    const rels: string[] = [];
    for (let i = 0; i < list.length; i++) {
      const p = list[i]!;
      rels.push(
        `<Relationship Id="rId${i + 1}" Type="${REL_PIVOT_TABLE}" Target="../pivotTables/pivotTable${p.cacheIdx}.xml"/>`,
      );
    }
    const relsXml =
      `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
      `<Relationships xmlns="${PKG_REL}">${rels.join("")}</Relationships>`;
    const pivotTablesFragment =
      `<pivotTables count="${list.length}">` +
      list
        .map((p, idx) => `<pivotTable name="${escapeXml(p.pivotTableName)}" r:id="rId${idx + 1}"/>`)
        .join("") +
      `</pivotTables>`;
    plans.push({ sheetIndex, relsXml, pivotTablesFragment });
  }
  return plans;
}
