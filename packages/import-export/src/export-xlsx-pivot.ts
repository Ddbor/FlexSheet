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
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition";
const REL_PIVOT_CACHE_RECORDS =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheRecords";
const REL_PIVOT_TABLE =
  "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable";

// Internal sentinels used as keys in indexMap for non-string values
const KEY_BOOL_FALSE = "\x00F";
const KEY_BOOL_TRUE = "\x00T";
const KEY_BLANK = "\x00M";

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

/**
 * sharedItems 构建结果。
 * - `hasListedItems=true`：items 已列出，pivotCacheRecords 中该字段必须使用 `<x v="idx"/>` 索引引用。
 * - `hasListedItems=false`：仅保存类型属性，records 使用直接值 `<n>/<s>/<b>/<m>`。
 */
interface SharedItemsResult {
  readonly xml: string;
  readonly keys: readonly string[];
  readonly hasListedItems: boolean;
  /** 仅当 hasListedItems=true 时有效；将单元格值映射为 sharedItems 索引。 */
  readonly indexMap: ReadonlyMap<string, number>;
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
 * Excel 要求 cacheField/@name 在透视缓存内唯一；重复列标题需加后缀。
 */
function dedupePivotFieldNames(names: readonly string[]): string[] {
  const used = new Set<string>();
  const out: string[] = [];
  for (const raw of names) {
    let base = raw.trim();
    if (base === "") {
      base = "Column";
    }
    let candidate = base;
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
 * 为单个字段构建 sharedItems XML。
 *
 * 规则（与 Excel 一致）：
 * - 含非整数小数的纯数值字段 → 仅输出属性（无条目列表），records 用直接值。
 * - 字符串、整数或布尔值字段 → 列出所有去重条目，records 用 `<x v="idx"/>` 索引。
 *
 * 当 sharedItems 有列出条目时，Excel 强制要求 pivotCacheRecords 用索引引用，
 * 否则报「文件损坏无法修复」。
 */
function buildSharedItemsXmlForField(
  src: Worksheet,
  def: WorksheetPivotTableDefinition,
  fieldCol: number,
): SharedItemsResult {
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
      if (v) hasTrue = true;
      else hasFalse = true;
      continue;
    }
    strings.add(String(v));
  }

  const hasBooleans = hasFalse || hasTrue;
  const hasNumbers = numbers.size > 0;
  const hasStrings = strings.size > 0;
  const numsSorted = [...numbers].sort((a, b) => a - b);
  const allIntegers = hasNumbers && numsSorted.every((v) => Number.isInteger(v));
  // 含非整数小数的纯数值字段：与 Excel 规范一致，只保留属性不列条目
  const hasNonIntegerDecimals = hasNumbers && !hasStrings && !hasBooleans && !allIntegers;

  const mixed =
    (hasNumbers && hasStrings) ||
    (hasNumbers && hasBooleans) ||
    (hasStrings && hasBooleans) ||
    (hasFalse && hasTrue);

  if (hasNonIntegerDecimals) {
    // 只输出属性，不列条目，records 继续用直接 <n v="..."/>
    const attrs: string[] = [];
    attrs.push('containsSemiMixedTypes="0"');
    attrs.push('containsString="0"');
    attrs.push('containsNumber="1"');
    if (hasBlank) attrs.push('containsBlank="1"');
    if (numsSorted.length > 0) {
      attrs.push(`minValue="${numsSorted[0]}"`);
      attrs.push(`maxValue="${numsSorted[numsSorted.length - 1]}"`);
    }
    return {
      xml: `<sharedItems ${attrs.join(" ")}/>`,
      keys: [],
      hasListedItems: false,
      indexMap: new Map(),
    };
  }

  // 列出所有条目，records 用 <x v="idx"/>
  const parts: string[] = [];
  const keys: string[] = [];
  const indexMap = new Map<string, number>();

  for (const nv of numsSorted) {
    indexMap.set(String(nv), parts.length);
    parts.push(`<n v="${String(nv)}"/>`);
    keys.push(String(nv));
  }
  const strsSorted = [...strings].sort((a, b) =>
    a.localeCompare(b, undefined, { sensitivity: "accent" }),
  );
  for (const s of strsSorted) {
    indexMap.set(s, parts.length);
    parts.push(`<s v="${escapeXml(s)}"/>`);
    keys.push(s);
  }
  if (hasFalse) {
    indexMap.set(KEY_BOOL_FALSE, parts.length);
    parts.push(`<b v="0"/>`);
    keys.push("FALSE");
  }
  if (hasTrue) {
    indexMap.set(KEY_BOOL_TRUE, parts.length);
    parts.push(`<b v="1"/>`);
    keys.push("TRUE");
  }
  if (hasBlank) {
    indexMap.set(KEY_BLANK, parts.length);
    parts.push(`<m/>`);
    keys.push("(空白)");
  }

  const attrs: string[] = [];
  attrs.push(mixed ? 'containsSemiMixedTypes="1"' : 'containsSemiMixedTypes="0"');
  if (!hasStrings) attrs.push('containsString="0"');
  if (hasNumbers) {
    attrs.push('containsNumber="1"');
    if (allIntegers) attrs.push('containsInteger="1"');
    attrs.push(`minValue="${numsSorted[0]}"`);
    attrs.push(`maxValue="${numsSorted[numsSorted.length - 1]}"`);
  }
  if (hasBlank) attrs.push('containsBlank="1"');
  if (mixed) attrs.push('containsMixedTypes="1"');
  const cnt = parts.length;
  if (cnt > 0) attrs.push(`count="${cnt}"`);

  const attrStr = attrs.length > 0 ? ` ${attrs.join(" ")}` : "";
  return {
    xml: `<sharedItems${attrStr}>${parts.join("")}</sharedItems>`,
    keys,
    hasListedItems: true,
    indexMap,
  };
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

/** 将单元格值转换为 indexMap 查找键。 */
function cellValueToIndexKey(value: CellScalar): string {
  if (value === null || value === "") return KEY_BLANK;
  if (typeof value === "number") return Number.isFinite(value) ? String(value) : KEY_BLANK;
  if (typeof value === "boolean") return value ? KEY_BOOL_TRUE : KEY_BOOL_FALSE;
  return String(value);
}

/**
 * 按数据源区域逐行写入缓存记录。
 * - 有 sharedItems 条目（hasListedItems=true）的字段：写 `<x v="idx"/>`（索引引用）。
 * - 无条目（hasListedItems=false，如小数值字段）的字段：写直接值 `<n>/<s>/<b>/<m>`。
 *
 * 混用会导致 Excel 报「文件损坏无法修复」。
 */
function buildPivotCacheRecordsXml(
  src: Worksheet,
  def: WorksheetPivotTableDefinition,
  sharedItemsResults: readonly SharedItemsResult[],
): { readonly xml: string; readonly recordCount: number } {
  const n = normalizeSelectionRange(def.sourceRange);
  const rStart = def.hasHeaders ? n.startRow + 1 : n.startRow;
  const rEnd = n.endRow;
  const rowsXml: string[] = [];
  let recordCount = 0;
  for (let r = rStart; r <= rEnd; r++) {
    const cells: string[] = [];
    for (let c = n.startCol; c <= n.endCol; c++) {
      const fi = c - n.startCol;
      const result = sharedItemsResults[fi];
      const v = src.getCell(r, c).value;
      if (result?.hasListedItems) {
        const key = cellValueToIndexKey(v);
        const idx = result.indexMap.get(key);
        cells.push(idx !== undefined ? `<x v="${idx}"/>` : `<m/>`);
      } else {
        cells.push(pivotRecordFragmentFromScalar(v));
      }
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
  sheetNames: readonly string[],
  fieldNames: readonly string[],
  sharedItemsResults: readonly SharedItemsResult[],
  recordCount: number,
): string {
  const n = normalizeSelectionRange(def.sourceRange);
  const sheetName = sheetNames[def.sourceSheetIndex] ?? `Sheet${def.sourceSheetIndex + 1}`;
  const ref = `${formatCellRef(n.startRow, n.startCol)}:${formatCellRef(n.endRow, n.endCol)}`;
  const fieldsXml = fieldNames
    .map((nm, fi) => {
      const shared = sharedItemsResults[fi]?.xml ?? `<sharedItems/>`;
      return `<cacheField name="${escapeXml(nm)}">${shared}</cacheField>`;
    })
    .join("");
  const saveData = recordCount > 0 ? 1 : 0;
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<pivotCacheDefinition xmlns="${SS_MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
    ` r:id="rId1" refreshOnLoad="1" recordCount="${recordCount}" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" saveData="${saveData}">` +
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
  sharedItemsResults: readonly SharedItemsResult[],
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
  const selectedKeysByFilterOffset = new Map<number, readonly string[]>();
  for (let i = 0; i < filterOffs.length; i++) {
    selectedKeysByFilterOffset.set(filterOffs[i]!, def.filterSelectedKeys[i] ?? []);
  }
  for (let fi = 0; fi < fieldNames.length; fi++) {
    if (dataFieldIndices.has(fi)) {
      pivotFieldsXml.push(`<pivotField dataField="1" showAll="0"/>`);
    } else if (filterOffSet.has(fi)) {
      const selected = selectedKeysByFilterOffset.get(fi) ?? [];
      const sharedKeys = sharedItemsResults[fi]?.keys ?? [];
      if (selected.length === 0 || sharedKeys.length === 0) {
        pivotFieldsXml.push(
          `<pivotField axis="axisPage" showAll="1"><items count="1"><item t="default"/></items></pivotField>`,
        );
      } else {
        const selectedSet = new Set(selected);
        const itemXml = sharedKeys
          .map((k, idx) => {
            return selectedSet.has(k) ? `<item x="${idx}"/>` : `<item x="${idx}" h="1"/>`;
          })
          .join("");
        const hasHit = sharedKeys.some((k) => selectedSet.has(k));
        if (!hasHit) {
          pivotFieldsXml.push(
            `<pivotField axis="axisPage" showAll="1"><items count="1"><item t="default"/></items></pivotField>`,
          );
        } else {
          pivotFieldsXml.push(
            `<pivotField axis="axisPage" showAll="0"><items count="${sharedKeys.length}">${itemXml}</items></pivotField>`,
          );
        }
      }
    } else if (rowOffSet.has(fi)) {
      const itemCount = sharedItemsResults[fi]?.keys.length ?? 0;
      const itemsXml =
        Array.from({ length: itemCount }, (_, idx) => `<item x="${idx}"/>`).join("") +
        `<item t="default"/>`;
      pivotFieldsXml.push(
        `<pivotField axis="axisRow" showAll="0"><items count="${itemCount + 1}">${itemsXml}</items></pivotField>`,
      );
    } else if (colOffSet.has(fi)) {
      const itemCount = sharedItemsResults[fi]?.keys.length ?? 0;
      const itemsXml =
        Array.from({ length: itemCount }, (_, idx) => `<item x="${idx}"/>`).join("") +
        `<item t="default"/>`;
      pivotFieldsXml.push(
        `<pivotField axis="axisCol" showAll="0"><items count="${itemCount + 1}">${itemsXml}</items></pivotField>`,
      );
    } else {
      pivotFieldsXml.push(`<pivotField/>`);
    }
  }

  const rowFieldsXml =
    rowOffs.length === 0
      ? ""
      : `<rowFields count="${rowOffs.length}">` +
        rowOffs.map((o) => `<field x="${o}"/>`).join("") +
        `</rowFields>`;

  let rowItemsXml: string;
  if (rowOffs.length === 0) {
    rowItemsXml = `<rowItems count="1"><i/></rowItems>`;
  } else if (rowOffs.length === 1) {
    const itemCount = sharedItemsResults[rowOffs[0]!]?.keys.length ?? 0;
    if (itemCount === 0) {
      rowItemsXml = `<rowItems count="1"><i/></rowItems>`;
    } else {
      const items = Array.from({ length: itemCount }, (_, i) => `<i><x v="${i}"/></i>`).join("");
      rowItemsXml = `<rowItems count="${itemCount + 1}">${items}<i t="grand"><x/></i></rowItems>`;
    }
  } else {
    // Multiple row fields: simplified single-entry placeholder (full combination enumeration is complex)
    rowItemsXml = `<rowItems count="1"><i>${rowOffs.map(() => `<x v="0"/>`).join("")}</i></rowItems>`;
  }

  let colFieldsXml = "";
  let colItemsXml = `<colItems count="1"><i/></colItems>`;
  if (colOffs.length === 1) {
    const itemCount = sharedItemsResults[colOffs[0]!]?.keys.length ?? 0;
    if (itemCount === 0) {
      colFieldsXml = `<colFields count="1"><field x="${colOffs[0]}"/></colFields>`;
      colItemsXml = `<colItems count="1"><i/></colItems>`;
    } else {
      colFieldsXml = `<colFields count="1"><field x="${colOffs[0]}"/></colFields>`;
      const items = Array.from({ length: itemCount }, (_, i) => `<i><x v="${i}"/></i>`).join("");
      colItemsXml = `<colItems count="${itemCount + 1}">${items}<i t="grand"><x/></i></colItems>`;
    }
  } else if (colOffs.length > 1) {
    colFieldsXml =
      `<colFields count="${colOffs.length}">` +
      colOffs.map((o) => `<field x="${o}"/>`).join("") +
      `</colFields>`;
    colItemsXml = `<colItems count="1"><i>${colOffs.map(() => `<x v="0"/>`).join("")}</i></colItems>`;
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
      )}" showDataAs="normal"/>`,
    );
  }
  const dataFieldsXml = `<dataFields count="${dataFieldParts.length}">${dataFieldParts.join("")}</dataFields>`;

  const pageFieldsXml =
    filterOffs.length === 0
      ? ""
      : `<pageFields count="${filterOffs.length}">${filterOffs
          .map((o) => `<pageField fld="${o}" hier="-1"/>`)
          .join("")}</pageFields>`;

  // firstHeaderRow: 0 基偏移至列标题行（等于筛选字段行数）
  // firstDataRow: 0 基偏移至第一数据行（筛选行数 + 1）
  const firstHeaderRow = filterOffs.length;
  const firstDataRow = filterOffs.length + 1;

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
    ` applyWidthHeightFormats="1"` +
    `${def.valueFieldsOnRows === true && valueSpecs.length > 1 && colOffs.length === 0 ? ` dataOnRows="1"` : ""}>` +
    `<location ref="${escapeXml(locRef)}" firstHeaderRow="${firstHeaderRow}" firstDataRow="${firstDataRow}" firstDataCol="1"/>` +
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
      const cacheId = cacheIdx;
      const fieldNames = collectFieldNames(src, def);
      if (fieldNames.length === 0) {
        continue;
      }
      const n = normalizeSelectionRange(def.sourceRange);
      const sharedItemsResults = fieldNames.map((_, fi) =>
        buildSharedItemsXmlForField(src, def, n.startCol + fi),
      );
      const { xml: pivotCacheRecordsXml, recordCount } = buildPivotCacheRecordsXml(
        src,
        def,
        sharedItemsResults,
      );
      const pivotCacheDefinitionXml = buildPivotCacheDefinitionXml(
        def,
        sheetNames,
        fieldNames,
        sharedItemsResults,
        recordCount,
      );
      const pivotCacheDefinitionRelsXml = buildPivotCacheDefinitionRelsXml(cacheIdx);
      const pivotTableXml = buildPivotTableDefinitionXml(
        def,
        fieldNames,
        sharedItemsResults,
        cacheId,
      );
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
    // Excel discovers pivot tables via .rels only; the <pivotTables> worksheet element is non-standard
    const pivotTablesFragment = "";
    plans.push({ sheetIndex, relsXml, pivotTablesFragment });
  }
  return plans;
}
