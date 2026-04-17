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
  /** Excel 通常要求 cache 部件通过关系引用 `pivotCacheRecords` 部件（可与 `recordCount=0` 一致）。 */
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
  return names;
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

function buildPivotCacheRecordsXml(): string {
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<pivotCacheRecords xmlns="${SS_MAIN}" count="0"/>`
  );
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
): string {
  const n = normalizeSelectionRange(def.sourceRange);
  const sheetName = sheetNames[def.sourceSheetIndex] ?? `Sheet${def.sourceSheetIndex + 1}`;
  const ref = `${formatCellRef(n.startRow, n.startCol)}:${formatCellRef(n.endRow, n.endCol)}`;
  const fieldsXml = fieldNames
    .map(
      (nm) => `<cacheField name="${escapeXml(nm)}"><sharedItems containsBlank="1"/></cacheField>`,
    )
    .join("");
  return (
    `<?xml version="1.0" encoding="UTF-8" standalone="yes"?>` +
    `<pivotCacheDefinition xmlns="${SS_MAIN}" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"` +
    ` refreshOnLoad="1" recordCount="0" createdVersion="6" refreshedVersion="6" minRefreshableVersion="3" saveData="0">` +
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

  const r0 = def.destinationRow;
  const c0 = def.destinationCol;
  const r1 = r0 + def.outputRowCount - 1;
  const c1 = c0 + def.outputColCount - 1;
  const locRef = `${formatCellRef(r0, c0)}:${formatCellRef(r1, c1)}`;

  const dataFieldIndices = new Set<number>();
  for (const v of valueSpecs) {
    dataFieldIndices.add(toOff(v.col));
    if (v.computed?.kind === "bucketRatio") {
      dataFieldIndices.add(toOff(v.computed.denominatorCol));
    }
  }
  const rowOffSet = new Set(rowOffs);
  const colOffSet = new Set(colOffs);

  const pivotFieldsXml: string[] = [];
  for (let fi = 0; fi < fieldNames.length; fi++) {
    if (dataFieldIndices.has(fi)) {
      pivotFieldsXml.push(`<pivotField dataField="1" showAll="0"/>`);
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
      const pivotCacheDefinitionXml = buildPivotCacheDefinitionXml(def, sheetNames, fieldNames);
      const pivotCacheRecordsXml = buildPivotCacheRecordsXml();
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
