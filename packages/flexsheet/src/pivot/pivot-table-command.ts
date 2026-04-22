import {
  buildUnconfiguredPivotPlaceholderMatrix,
  getPivotValueFieldCaption,
  isUnconfiguredPivotDefinition,
  normalizeSelectionRange,
  pivotLayoutStartRow,
  Worksheet,
  type CellScalar,
  type CellStyle,
  type ICommand,
  type PivotAggregateKind,
  type PivotValueFieldSpec,
  type SelectionRange,
  type Workbook,
  type WorksheetPivotTableDefinition,
} from "@flexsheet/core";

export type { PivotAggregateKind, PivotValueFieldSpec } from "@flexsheet/core";

export interface PivotTableDestinationNewSheet {
  readonly kind: "newSheet";
  readonly preferredName?: string;
}

export interface PivotTableDestinationExistingSheet {
  readonly kind: "existingSheet";
  /** 透视表输出写入的工作表；省略时与数据源工作表相同。 */
  readonly targetSheet?: Worksheet;
  readonly startRow: number;
  readonly startCol: number;
}

export type PivotTableDestination =
  | PivotTableDestinationNewSheet
  | PivotTableDestinationExistingSheet;

export interface PivotTableBuildOptions {
  readonly sourceRange: SelectionRange;
  readonly hasHeaders: boolean;
  readonly rowFieldCols: readonly number[];
  readonly columnFieldCols: readonly number[];
  readonly filterFieldCols: readonly number[];
  /** 与 `filterFieldCols` 等长；省略或空子数组表示该筛选项不限制。 */
  readonly filterSelectedKeys?: readonly (readonly string[])[];
  /** 多值且无列字段时：度量在行展开（Excel「数值」在行）。 */
  readonly valueFieldsOnRows?: boolean;
  readonly valueFields: readonly PivotValueFieldSpec[];
  readonly destination: PivotTableDestination;
  /**
   * 与 `isUnconfiguredPivotDefinition` 一致时：按给定尺寸绘制「尚未放置字段」占位（Excel 导入空透视 / 刷新）。
   */
  readonly unconfiguredPlaceholder?: {
    readonly rowCount: number;
    readonly colCount: number;
    readonly title: string;
  };
}

interface PivotFieldItem {
  readonly col: number;
  readonly name: string;
}

interface PivotAccumulator {
  countAll: number;
  countNumeric: number;
  sum: number;
  min: number;
  max: number;
}

/** 单个值字段对应的累加状态（普通汇总 / 比率分子分母 / 占比分子）。 */
type PivotFieldAggState =
  | { readonly mode: "plain"; readonly acc: PivotAccumulator }
  | { readonly mode: "ratio"; readonly num: PivotAccumulator; readonly den: PivotAccumulator }
  | { readonly mode: "share"; readonly acc: PivotAccumulator };

interface PivotRenderOutput {
  readonly rowCount: number;
  readonly colCount: number;
  readonly values: readonly (readonly CellScalar[])[];
  readonly styles: readonly (readonly (CellStyle | null)[])[];
}

interface CellSnapshot {
  readonly row: number;
  readonly col: number;
  readonly formula: string | null;
  readonly value: CellScalar;
  readonly style: CellStyle | null;
}

const KEY_SEP = "\x1e";

function cloneStyle(style: CellStyle | null): CellStyle | null {
  return style === null ? null : { ...style };
}

function nextDefaultPivotSheetName(workbook: Workbook): string {
  let max = 0;
  for (let i = 0; i < workbook.sheetCount; i++) {
    const sheet = workbook.getSheet(i);
    if (sheet === undefined) {
      continue;
    }
    const m = /^数据透视表(\d+)?$/.exec(sheet.name);
    if (m === null) {
      continue;
    }
    if (m[1] === undefined) {
      max = Math.max(max, 1);
      continue;
    }
    const idx = Number.parseInt(m[1], 10);
    if (Number.isFinite(idx)) {
      max = Math.max(max, idx);
    }
  }
  return max === 0 ? "数据透视表1" : `数据透视表${max + 1}`;
}

/** 与透视行/列键、筛选项一致的单元格展示键。 */
export function pivotFilterKeyFromCellValue(value: CellScalar): string {
  return toPivotKey(value);
}

function toPivotKey(value: CellScalar): string {
  if (value === null) {
    return "(空白)";
  }
  if (typeof value === "string") {
    const trimmed = value.trim();
    return trimmed.length > 0 ? trimmed : "(空白)";
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  if (Number.isFinite(value)) {
    return String(value);
  }
  return "(空白)";
}

function toFieldDisplayName(value: CellScalar, fallback: string): string {
  if (typeof value === "string") {
    const trimmed = value.trim();
    if (trimmed.length > 0) {
      return trimmed;
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

function createAccumulator(): PivotAccumulator {
  return {
    countAll: 0,
    countNumeric: 0,
    sum: 0,
    min: Number.POSITIVE_INFINITY,
    max: Number.NEGATIVE_INFINITY,
  };
}

function pushAggregate(acc: PivotAccumulator, value: CellScalar): void {
  const isBlank = value === null || (typeof value === "string" && value.trim().length === 0);
  if (!isBlank) {
    acc.countAll++;
  }
  if (typeof value !== "number" || !Number.isFinite(value)) {
    return;
  }
  acc.countNumeric++;
  acc.sum += value;
  acc.min = Math.min(acc.min, value);
  acc.max = Math.max(acc.max, value);
}

function reduceAggregate(kind: PivotAggregateKind, acc: PivotAccumulator): CellScalar {
  switch (kind) {
    case "count":
      return acc.countAll;
    case "sum":
      return acc.countNumeric > 0 ? acc.sum : null;
    case "average":
      return acc.countNumeric > 0 ? acc.sum / acc.countNumeric : null;
    case "max":
      return acc.countNumeric > 0 ? acc.max : null;
    case "min":
      return acc.countNumeric > 0 ? acc.min : null;
    default:
      return null;
  }
}

function collectFields(
  sheet: Worksheet,
  range: SelectionRange,
  hasHeaders: boolean,
): readonly PivotFieldItem[] {
  const n = normalizeSelectionRange(range);
  const list: PivotFieldItem[] = [];
  for (let c = n.startCol; c <= n.endCol; c++) {
    const fallback = `列${c - n.startCol + 1}`;
    const name = hasHeaders
      ? toFieldDisplayName(sheet.getCell(n.startRow, c).value, fallback)
      : fallback;
    list.push({ col: c, name });
  }
  return list;
}

function rowCompositeKey(sheet: Worksheet, row: number, cols: readonly number[]): string {
  return cols.map((c) => toPivotKey(sheet.getCell(row, c).value)).join(KEY_SEP);
}

function colCompositeKey(sheet: Worksheet, row: number, cols: readonly number[]): string {
  return cols.map((c) => toPivotKey(sheet.getCell(row, c).value)).join(KEY_SEP);
}

function mergeAccumulator(target: PivotAccumulator, from: PivotAccumulator): void {
  target.countAll += from.countAll;
  target.countNumeric += from.countNumeric;
  target.sum += from.sum;
  target.min = Math.min(target.min, from.min);
  target.max = Math.max(target.max, from.max);
}

function createFieldAggState(vf: PivotValueFieldSpec): PivotFieldAggState {
  if (vf.computed?.kind === "bucketRatio") {
    return { mode: "ratio", num: createAccumulator(), den: createAccumulator() };
  }
  if (vf.computed?.kind === "shareOfGrandTotal") {
    return { mode: "share", acc: createAccumulator() };
  }
  return { mode: "plain", acc: createAccumulator() };
}

function pushFieldAgg(
  vf: PivotValueFieldSpec,
  state: PivotFieldAggState,
  sheet: Worksheet,
  row: number,
): void {
  if (state.mode === "plain") {
    pushAggregate(state.acc, sheet.getCell(row, vf.col).value);
    return;
  }
  if (state.mode === "ratio" && vf.computed?.kind === "bucketRatio") {
    pushAggregate(state.num, sheet.getCell(row, vf.col).value);
    pushAggregate(state.den, sheet.getCell(row, vf.computed.denominatorCol).value);
    return;
  }
  if (state.mode === "share") {
    pushAggregate(state.acc, sheet.getCell(row, vf.col).value);
  }
}

function mergeFieldAggState(target: PivotFieldAggState, from: PivotFieldAggState): void {
  if (target.mode === "plain" && from.mode === "plain") {
    mergeAccumulator(target.acc, from.acc);
    return;
  }
  if (target.mode === "ratio" && from.mode === "ratio") {
    mergeAccumulator(target.num, from.num);
    mergeAccumulator(target.den, from.den);
    return;
  }
  if (target.mode === "share" && from.mode === "share") {
    mergeAccumulator(target.acc, from.acc);
  }
}

function reduceFieldAgg(
  vf: PivotValueFieldSpec,
  state: PivotFieldAggState,
  globalShareDenom: ReadonlyMap<number, number>,
): CellScalar {
  if (state.mode === "plain") {
    return reduceAggregate(vf.aggregate, state.acc);
  }
  if (state.mode === "ratio") {
    if (state.den.countNumeric === 0 || state.den.sum === 0) {
      return null;
    }
    if (state.num.countNumeric === 0) {
      return null;
    }
    return state.num.sum / state.den.sum;
  }
  const g = globalShareDenom.get(vf.col) ?? 0;
  if (g <= 0 || state.acc.countNumeric === 0) {
    return null;
  }
  return state.acc.sum / g;
}

function computeGlobalShareDenominators(
  sheet: Worksheet,
  rStart: number,
  rEnd: number,
  shareCols: ReadonlySet<number>,
  rowInclude: (r: number) => boolean,
): Map<number, number> {
  const out = new Map<number, number>();
  for (const c of shareCols) {
    out.set(c, 0);
  }
  for (let r = rStart; r <= rEnd; r++) {
    if (!rowInclude(r)) {
      continue;
    }
    for (const c of shareCols) {
      const v = sheet.getCell(r, c).value;
      if (typeof v === "number" && Number.isFinite(v)) {
        out.set(c, (out.get(c) ?? 0) + v);
      }
    }
  }
  return out;
}

export function normalizeFilterSelectedKeys(
  filterFieldCols: readonly number[],
  input: readonly (readonly string[])[] | undefined,
): readonly (readonly string[])[] {
  const out: string[][] = [];
  for (let i = 0; i < filterFieldCols.length; i++) {
    const row = input?.[i];
    out.push(row !== undefined && row.length > 0 ? [...row] : []);
  }
  return out;
}

/** 布局变更时保留仍存在的筛选列上的选中项。 */
export function mergePivotFilterSelections(
  oldCols: readonly number[],
  oldKeys: readonly (readonly string[])[] | undefined,
  newCols: readonly number[],
  newKeysFromLayout: readonly (readonly string[])[] | undefined,
): readonly (readonly string[])[] {
  // oldKeys 缺省（旧数据）视为与 filterFieldCols 等长的空限制列表
  const oldMap = new Map<number, readonly string[]>();
  const ok = oldKeys ?? [];
  for (let i = 0; i < oldCols.length; i++) {
    oldMap.set(oldCols[i]!, ok[i] ?? []);
  }
  const out: string[][] = [];
  for (let i = 0; i < newCols.length; i++) {
    const c = newCols[i]!;
    const fromLayout = newKeysFromLayout?.[i];
    if (fromLayout !== undefined) {
      out.push([...fromLayout]);
      continue;
    }
    const prev = oldMap.get(c);
    out.push(prev !== undefined ? [...prev] : []);
  }
  return out;
}

function rowPassesPivotFilters(
  sheet: Worksheet,
  row: number,
  filterFieldCols: readonly number[],
  filterSelectedKeys: readonly (readonly string[])[],
): boolean {
  for (let i = 0; i < filterFieldCols.length; i++) {
    const allowed = filterSelectedKeys[i];
    if (allowed === undefined || allowed.length === 0) {
      continue;
    }
    const fc = filterFieldCols[i]!;
    const key = toPivotKey(sheet.getCell(row, fc).value);
    if (!allowed.includes(key)) {
      return false;
    }
  }
  return true;
}

function formatPivotFilterCaption(selectedKeys: readonly string[]): string {
  if (selectedKeys.length === 0) {
    return "(全部)";
  }
  if (selectedKeys.length === 1) {
    return selectedKeys[0]!;
  }
  return "(多项)";
}

/** 数据源列上去重后的透视键（用于筛选面板列表）。 */
export function collectPivotFilterDistinctKeys(
  sheet: Worksheet,
  rStart: number,
  rEnd: number,
  col: number,
): string[] {
  const set = new Set<string>();
  for (let r = rStart; r <= rEnd; r++) {
    set.add(toPivotKey(sheet.getCell(r, col).value));
  }
  return [...set].sort((a, b) => a.localeCompare(b, "zh-Hans-CN"));
}

export function dataFieldCaption(
  spec: PivotValueFieldSpec,
  fieldName: string,
  denominatorFieldName?: string,
): string {
  return getPivotValueFieldCaption(spec, fieldName, denominatorFieldName);
}

function styleForValueAggregate(agg: PivotAggregateKind): CellStyle | null {
  if (agg === "count") {
    return { numberFormat: "#,##0" };
  }
  return { numberFormat: "#,##0.00" };
}

function styleForValueField(vf: PivotValueFieldSpec): CellStyle | null {
  if (vf.computed?.kind === "bucketRatio" || vf.computed?.kind === "shareOfGrandTotal") {
    return { numberFormat: "0.00%" };
  }
  return styleForValueAggregate(vf.aggregate);
}

function dedupeColsWithinZone(cols: number[]): void {
  const seen = new Set<number>();
  for (let i = cols.length - 1; i >= 0; i--) {
    const c = cols[i]!;
    if (seen.has(c)) {
      cols.splice(i, 1);
    } else {
      seen.add(c);
    }
  }
}

function dedupeValueFieldsByColInPlace(valueFields: PivotValueFieldSpec[]): void {
  const seen = new Set<number>();
  for (let i = valueFields.length - 1; i >= 0; i--) {
    const c = valueFields[i]!.col;
    if (seen.has(c)) {
      valueFields.splice(i, 1);
    } else {
      seen.add(c);
    }
  }
}

/** 同一列不可同时出现在筛选/列/行/值；优先级：筛选 > 列 > 行 > 值（从值区移除重复项）。 */
function dedupePivotColumnAcrossZones(
  filterFieldCols: number[],
  columnFieldCols: number[],
  rowFieldCols: number[],
  valueFields: PivotValueFieldSpec[],
): void {
  dedupeColsWithinZone(filterFieldCols);
  dedupeColsWithinZone(columnFieldCols);
  dedupeColsWithinZone(rowFieldCols);
  dedupeValueFieldsByColInPlace(valueFields);
  const used = new Set<number>();
  for (const c of filterFieldCols) {
    used.add(c);
  }
  for (let i = columnFieldCols.length - 1; i >= 0; i--) {
    const c = columnFieldCols[i]!;
    if (used.has(c)) {
      columnFieldCols.splice(i, 1);
    } else {
      used.add(c);
    }
  }
  for (let i = rowFieldCols.length - 1; i >= 0; i--) {
    const c = rowFieldCols[i]!;
    if (used.has(c)) {
      rowFieldCols.splice(i, 1);
    } else {
      used.add(c);
    }
  }
  for (let i = valueFields.length - 1; i >= 0; i--) {
    const c = valueFields[i]!.col;
    if (used.has(c)) {
      valueFields.splice(i, 1);
    } else {
      used.add(c);
    }
  }
}

export function buildPivotRender(
  sheet: Worksheet,
  options: PivotTableBuildOptions,
): PivotRenderOutput {
  const n = normalizeSelectionRange(options.sourceRange);
  const fields = collectFields(sheet, n, options.hasHeaders);
  const fieldNameByCol = new Map<number, string>(fields.map((it) => [it.col, it.name]));

  const rowCols = [...options.rowFieldCols];
  const colDimCols = [...options.columnFieldCols];
  let valueFields = [...options.valueFields.map((v) => ({ ...v }))];
  const filterFieldCols = [...options.filterFieldCols];
  dedupePivotColumnAcrossZones(filterFieldCols, colDimCols, rowCols, valueFields);
  const filterSelectedKeysNorm = normalizeFilterSelectedKeys(filterFieldCols, options.filterSelectedKeys);

  const ph = options.unconfiguredPlaceholder;
  if (
    ph !== undefined &&
    rowCols.length === 0 &&
    colDimCols.length === 0 &&
    valueFields.length === 0 &&
    filterFieldCols.length === 0
  ) {
    const m = buildUnconfiguredPivotPlaceholderMatrix(ph.rowCount, ph.colCount, ph.title);
    return {
      rowCount: m.rowCount,
      colCount: m.colCount,
      values: m.values,
      styles: m.styles,
    };
  }

  if (rowCols.length === 0 || valueFields.length === 0) {
    return {
      rowCount: 1,
      colCount: 1,
      values: [["(无效布局)"]],
      styles: [[null]],
    };
  }
  if (colDimCols.length > 0 && valueFields.length > 1) {
    valueFields = [valueFields[0]!];
  }

  const rStart = options.hasHeaders ? n.startRow + 1 : n.startRow;
  const rEnd = n.endRow;

  const rowInclude = (r: number): boolean =>
    rowPassesPivotFilters(sheet, r, filterFieldCols, filterSelectedKeysNorm);

  const shareCols = new Set<number>();
  for (const vf of valueFields) {
    if (vf.computed?.kind === "shareOfGrandTotal") {
      shareCols.add(vf.col);
    }
  }
  const globalShareDenom = computeGlobalShareDenominators(sheet, rStart, rEnd, shareCols, rowInclude);

  const matrix = new Map<string, Map<string, PivotFieldAggState[]>>();
  const rowKeySet = new Set<string>();
  const colKeySet = new Set<string>();

  for (let r = rStart; r <= rEnd; r++) {
    if (!rowInclude(r)) {
      continue;
    }
    const rowKey = rowCompositeKey(sheet, r, rowCols);
    const colKey = colDimCols.length === 0 ? "__single__" : colCompositeKey(sheet, r, colDimCols);
    rowKeySet.add(rowKey);
    colKeySet.add(colKey);
    let rowMap = matrix.get(rowKey);
    if (rowMap === undefined) {
      rowMap = new Map<string, PivotFieldAggState[]>();
      matrix.set(rowKey, rowMap);
    }
    let accs = rowMap.get(colKey);
    if (accs === undefined) {
      accs = valueFields.map((vf) => createFieldAggState(vf));
      rowMap.set(colKey, accs);
    }
    for (let vi = 0; vi < valueFields.length; vi++) {
      const vf = valueFields[vi]!;
      pushFieldAgg(vf, accs[vi]!, sheet, r);
    }
  }

  const rowKeys = [...rowKeySet].sort((a, b) => a.localeCompare(b, "zh-Hans-CN"));
  const colKeys = [...colKeySet].sort((a, b) => a.localeCompare(b, "zh-Hans-CN"));

  const showColDim = colDimCols.length > 0;
  const singleValueNoColDim = !showColDim && valueFields.length === 1;
  const multiValueNoColDim = !showColDim && valueFields.length > 1;
  const multiValueTall =
    multiValueNoColDim && options.valueFieldsOnRows === true;
  const multiValueWide = multiValueNoColDim && !multiValueTall;

  const colKeysOrdered = colKeys.length === 0 ? ["__single__"] : colKeys;
  let dataColCount: number;
  if (showColDim) {
    dataColCount = colKeysOrdered.length;
  } else if (multiValueTall) {
    dataColCount = 1;
  } else {
    dataColCount = valueFields.length;
  }
  const rowTotalColCount = singleValueNoColDim ? 1 : 0;
  const tableCols = 1 + dataColCount + rowTotalColCount;
  const innerBodyRowCount = multiValueTall ? rowKeys.length * valueFields.length : rowKeys.length;
  const grandRowCount = multiValueTall ? valueFields.length : 1;
  const innerTableRows = 1 + innerBodyRowCount + grandRowCount;
  const top = filterFieldCols.length;
  const totalRows = top + innerTableRows;

  const values: CellScalar[][] = Array.from({ length: totalRows }, () =>
    Array.from<CellScalar>({ length: tableCols }).fill(null),
  );
  const styles: (CellStyle | null)[][] = Array.from({ length: totalRows }, () =>
    Array.from<CellStyle | null>({ length: tableCols }).fill(null),
  );

  const filterLabelStyle: CellStyle = {
    bold: true,
    fillArgb: "FFF2F2F2",
  };
  const filterValueStyle: CellStyle = {
    fillArgb: "FFFFFFFF",
  };

  for (let fi = 0; fi < top; fi++) {
    const fc = filterFieldCols[fi]!;
    values[fi][0] = fieldNameByCol.get(fc) ?? `列${fc - n.startCol + 1}`;
    values[fi][1] = formatPivotFilterCaption(filterSelectedKeysNorm[fi] ?? []);
    styles[fi][0] = filterLabelStyle;
    styles[fi][1] = filterValueStyle;
  }

  const headStyle: CellStyle = {
    bold: true,
    fillArgb: "FFE2EFDA",
    hAlign: "center",
  };
  const totalStyle: CellStyle = {
    bold: true,
    fillArgb: "FFF2F2F2",
  };

  const rowHeaderLabel =
    rowCols.length === 1 ? (fieldNameByCol.get(rowCols[0]!) ?? "行") : "行标签";
  const colDimFirstName = colDimCols.length > 0 ? (fieldNameByCol.get(colDimCols[0]!) ?? "列") : "";

  const headerRow = top;
  values[headerRow][0] =
    showColDim && colDimCols.length > 1
      ? `${rowHeaderLabel}\\${colDimCols.map((c) => fieldNameByCol.get(c) ?? "").join("\\")}`
      : showColDim
        ? `${rowHeaderLabel}\\${colDimFirstName}`
        : rowHeaderLabel;
  styles[headerRow][0] = headStyle;

  if (showColDim) {
    const vf0 = valueFields[0]!;
    const vfName = fieldNameByCol.get(vf0.col) ?? "值";
    for (let ci = 0; ci < colKeysOrdered.length; ci++) {
      const ck = colKeysOrdered[ci]!;
      values[headerRow][ci + 1] = ck === "__single__" ? vfName : ck;
      styles[headerRow][ci + 1] = headStyle;
    }
    values[headerRow][tableCols - 1] = "总计";
    styles[headerRow][tableCols - 1] = headStyle;
  } else if (multiValueWide) {
    for (let vi = 0; vi < valueFields.length; vi++) {
      const vf = valueFields[vi]!;
      const nm = fieldNameByCol.get(vf.col) ?? "值";
      const denNm =
        vf.computed?.kind === "bucketRatio"
          ? (fieldNameByCol.get(vf.computed.denominatorCol) ?? "")
          : undefined;
      values[headerRow][vi + 1] = getPivotValueFieldCaption(vf, nm, denNm);
      styles[headerRow][vi + 1] = headStyle;
    }
  } else if (multiValueTall) {
    values[headerRow][1] = "汇总";
    styles[headerRow][1] = headStyle;
  } else {
    const vf0 = valueFields[0]!;
    const vfName = fieldNameByCol.get(vf0.col) ?? "值";
    const denNm =
      vf0.computed?.kind === "bucketRatio"
        ? (fieldNameByCol.get(vf0.computed.denominatorCol) ?? "")
        : undefined;
    values[headerRow][1] = getPivotValueFieldCaption(vf0, vfName, denNm);
    styles[headerRow][1] = headStyle;
    values[headerRow][2] = "总计";
    styles[headerRow][2] = headStyle;
  }

  const vf0ForColDim = valueFields[0]!;
  const colTotals: PivotFieldAggState[] = colKeysOrdered.map(() =>
    createFieldAggState(vf0ForColDim),
  );
  const allTotal: PivotFieldAggState[] = valueFields.map((vf) => createFieldAggState(vf));

  if (multiValueTall) {
    const colKeySingle = colKeysOrdered[0]!;
    let dst = headerRow + 1;
    for (let ri = 0; ri < rowKeys.length; ri++) {
      const rowKey = rowKeys[ri]!;
      const rowKeyText = rowKey.split(KEY_SEP).join(" ");
      const rowMap = matrix.get(rowKey) ?? new Map<string, PivotFieldAggState[]>();
      for (let vi = 0; vi < valueFields.length; vi++) {
        const vf = valueFields[vi]!;
        const nm = fieldNameByCol.get(vf.col) ?? "值";
        const denNm =
          vf.computed?.kind === "bucketRatio"
            ? (fieldNameByCol.get(vf.computed.denominatorCol) ?? "")
            : undefined;
        const caption = getPivotValueFieldCaption(vf, nm, denNm);
        values[dst][0] = vi === 0 ? rowKeyText : `　${caption}`;
        const accs = rowMap.get(colKeySingle);
        const acc = accs?.[vi];
        const out = acc === undefined ? null : reduceFieldAgg(vf, acc, globalShareDenom);
        values[dst][1] = out;
        styles[dst][0] = null;
        styles[dst][1] = styleForValueField(vf);
        if (acc !== undefined) {
          mergeFieldAggState(allTotal[vi]!, acc);
        }
        dst++;
      }
    }
    let tRow = dst;
    for (let vi = 0; vi < valueFields.length; vi++) {
      const vf = valueFields[vi]!;
      const nm = fieldNameByCol.get(vf.col) ?? "值";
      const denNm =
        vf.computed?.kind === "bucketRatio"
          ? (fieldNameByCol.get(vf.computed.denominatorCol) ?? "")
          : undefined;
      const cap = getPivotValueFieldCaption(vf, nm, denNm);
      values[tRow][0] = `总计 ${cap}`;
      values[tRow][1] = reduceFieldAgg(vf, allTotal[vi]!, globalShareDenom);
      styles[tRow][0] = totalStyle;
      styles[tRow][1] = totalStyle;
      tRow++;
    }
  } else {
    for (let ri = 0; ri < rowKeys.length; ri++) {
      const rowKey = rowKeys[ri]!;
      const rowMap = matrix.get(rowKey) ?? new Map<string, PivotFieldAggState[]>();
      const dst = headerRow + 1 + ri;
      values[dst][0] = rowKey.split(KEY_SEP).join(" ");
      styles[dst][0] = null;
      const rowTotalsAcc: PivotFieldAggState[] = valueFields.map((vf) => createFieldAggState(vf));

      for (let ci = 0; ci < colKeysOrdered.length; ci++) {
        const colKey = colKeysOrdered[ci]!;
        const accs = rowMap.get(colKey);
        if (showColDim) {
          const vf0 = valueFields[0]!;
          const acc = accs?.[0];
          const out = acc === undefined ? null : reduceFieldAgg(vf0, acc, globalShareDenom);
          values[dst][ci + 1] = out;
          styles[dst][ci + 1] = styleForValueField(vf0);
          if (acc !== undefined) {
            mergeFieldAggState(rowTotalsAcc[0]!, acc);
            mergeFieldAggState(colTotals[ci]!, acc);
            mergeFieldAggState(allTotal[0]!, acc);
          }
        } else if (multiValueWide) {
          for (let vi = 0; vi < valueFields.length; vi++) {
            const vf = valueFields[vi]!;
            const acc = accs?.[vi];
            const out = acc === undefined ? null : reduceFieldAgg(vf, acc, globalShareDenom);
            values[dst][vi + 1] = out;
            styles[dst][vi + 1] = styleForValueField(vf);
            if (acc !== undefined) {
              mergeFieldAggState(rowTotalsAcc[vi]!, acc);
              mergeFieldAggState(allTotal[vi]!, acc);
            }
          }
        } else {
          const vf0 = valueFields[0]!;
          const acc = accs?.[0];
          const out = acc === undefined ? null : reduceFieldAgg(vf0, acc, globalShareDenom);
          values[dst][1] = out;
          styles[dst][1] = styleForValueField(vf0);
          if (acc !== undefined) {
            mergeFieldAggState(rowTotalsAcc[0]!, acc);
            mergeFieldAggState(colTotals[ci]!, acc);
            mergeFieldAggState(allTotal[0]!, acc);
          }
          values[dst][2] = reduceFieldAgg(vf0, rowTotalsAcc[0]!, globalShareDenom);
          styles[dst][2] = totalStyle;
        }
      }

      if (showColDim) {
        const vf0 = valueFields[0]!;
        values[dst][tableCols - 1] = reduceFieldAgg(vf0, rowTotalsAcc[0]!, globalShareDenom);
        styles[dst][tableCols - 1] = totalStyle;
      }
    }

    const totalRowIdx = headerRow + rowKeys.length + 1;
    values[totalRowIdx][0] = "总计";
    styles[totalRowIdx][0] = totalStyle;
    if (showColDim) {
      const vf0 = valueFields[0]!;
      for (let ci = 0; ci < colKeysOrdered.length; ci++) {
        const colTotal = colTotals[ci]!;
        values[totalRowIdx][ci + 1] = reduceFieldAgg(vf0, colTotal, globalShareDenom);
        styles[totalRowIdx][ci + 1] = totalStyle;
      }
      values[totalRowIdx][tableCols - 1] = reduceFieldAgg(vf0, allTotal[0]!, globalShareDenom);
      styles[totalRowIdx][tableCols - 1] = totalStyle;
    } else if (multiValueWide) {
      for (let vi = 0; vi < valueFields.length; vi++) {
        const vf = valueFields[vi]!;
        values[totalRowIdx][vi + 1] = reduceFieldAgg(vf, allTotal[vi]!, globalShareDenom);
        styles[totalRowIdx][vi + 1] = totalStyle;
      }
    } else {
      const vf0 = valueFields[0]!;
      values[totalRowIdx][1] = reduceFieldAgg(vf0, allTotal[0]!, globalShareDenom);
      styles[totalRowIdx][1] = totalStyle;
      values[totalRowIdx][2] = reduceFieldAgg(vf0, allTotal[0]!, globalShareDenom);
      styles[totalRowIdx][2] = totalStyle;
    }
  }

  return {
    rowCount: totalRows,
    colCount: tableCols,
    values,
    styles,
  };
}

function writePivotResult(
  sheet: Worksheet,
  startRow: number,
  startCol: number,
  output: PivotRenderOutput,
): void {
  const needRows = startRow + output.rowCount;
  const needCols = startCol + output.colCount;
  if (needRows > sheet.rowCount || needCols > sheet.colCount) {
    sheet.setGridSize(Math.max(sheet.rowCount, needRows), Math.max(sheet.colCount, needCols));
  }
  sheet.batch(() => {
    for (let r = 0; r < output.rowCount; r++) {
      for (let c = 0; c < output.colCount; c++) {
        const rr = startRow + r;
        const cc = startCol + c;
        sheet.setCellLiteral(rr, cc, output.values[r]?.[c] ?? null);
        sheet.setCellStyle(rr, cc, cloneStyle(output.styles[r]?.[c] ?? null));
      }
    }
  });
}

/** 在曾用于更大透视输出的矩形内擦除收缩后多出的单元格，避免旧数据残留。 */
function clearPivotOutputOverflow(
  sheet: Worksheet,
  destRow: number,
  destCol: number,
  snapRows: number,
  snapCols: number,
  outRows: number,
  outCols: number,
): void {
  if (snapRows <= outRows && snapCols <= outCols) {
    return;
  }
  sheet.batch(() => {
    for (let r = destRow; r < destRow + snapRows; r++) {
      for (let c = destCol; c < destCol + snapCols; c++) {
        if (r >= destRow + outRows || c >= destCol + outCols) {
          sheet.setCellLiteral(r, c, null);
          sheet.setCellStyle(r, c, null);
        }
      }
    }
  });
}

function captureSnapshots(
  sheet: Worksheet,
  startRow: number,
  startCol: number,
  rowCount: number,
  colCount: number,
): readonly CellSnapshot[] {
  const out: CellSnapshot[] = [];
  for (let r = 0; r < rowCount; r++) {
    for (let c = 0; c < colCount; c++) {
      const rr = startRow + r;
      const cc = startCol + c;
      if (rr < 0 || cc < 0 || rr >= sheet.rowCount || cc >= sheet.colCount) {
        out.push({ row: rr, col: cc, formula: null, value: null, style: null });
        continue;
      }
      const cell = sheet.getCell(rr, cc);
      out.push({
        row: rr,
        col: cc,
        formula: cell.formula,
        value: cell.value,
        style: cloneStyle(cell.style),
      });
    }
  }
  return out;
}

function restoreSnapshots(sheet: Worksheet, snapshots: readonly CellSnapshot[]): void {
  sheet.batch(() => {
    for (const snap of snapshots) {
      if (snap.formula !== null) {
        sheet.setCellFormula(snap.row, snap.col, snap.formula);
      } else {
        sheet.setCellLiteral(snap.row, snap.col, snap.value);
      }
      sheet.setCellStyle(snap.row, snap.col, cloneStyle(snap.style));
    }
  });
}

function findSheetIndex(workbook: Workbook, target: Worksheet): number {
  for (let i = 0; i < workbook.sheetCount; i++) {
    if (workbook.getSheet(i) === target) {
      return i;
    }
  }
  return -1;
}

function nextPivotTableExcelName(workbook: Workbook): string {
  let max = 0;
  for (let i = 0; i < workbook.sheetCount; i++) {
    const sh = workbook.getSheet(i);
    if (sh === undefined) {
      continue;
    }
    for (const p of sh.getPivotTableDefinitionsSnapshot()) {
      const m = /^PivotTable(\d+)$/.exec(p.name);
      if (m !== null) {
        const n = Number.parseInt(m[1]!, 10);
        if (Number.isFinite(n)) {
          max = Math.max(max, n);
        }
      }
    }
  }
  return `PivotTable${max + 1}`;
}

function newPivotDefId(): string {
  return `pvt-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
}

function clonePivotDef(def: WorksheetPivotTableDefinition): WorksheetPivotTableDefinition {
  return {
    ...def,
    sourceRange: { ...def.sourceRange },
    rowFieldCols: [...def.rowFieldCols],
    columnFieldCols: [...def.columnFieldCols],
    filterFieldCols: [...def.filterFieldCols],
    filterSelectedKeys: (def.filterSelectedKeys ?? []).map((k) => [...k]),
    valueFields: def.valueFields.map((v) => ({ ...v })),
    valueFieldsOnRows: def.valueFieldsOnRows,
  };
}

/**
 * 按当前已落地的透视定义重算并覆写输出区域（不进入撤销栈）。
 * 返回 `true` 表示找到并刷新成功，`false` 表示定义或源表不存在。
 */
export function refreshPivotTableDefinition(
  workbook: Workbook,
  pivotSheet: Worksheet,
  pivotDefId: string,
): boolean {
  const current = pivotSheet.getPivotTableDefinitionsSnapshot().find((d) => d.id === pivotDefId);
  if (current === undefined) {
    return false;
  }
  const sourceSheet = workbook.getSheet(current.sourceSheetIndex);
  if (sourceSheet === undefined) {
    return false;
  }
  const layoutTop = pivotLayoutStartRow(current);
  const unconfigured =
    isUnconfiguredPivotDefinition(current) === true
      ? {
          rowCount: current.outputRowCount,
          colCount: current.outputColCount,
          title: current.name,
        }
      : undefined;
  const out = buildPivotRender(sourceSheet, {
    sourceRange: current.sourceRange,
    hasHeaders: current.hasHeaders,
    rowFieldCols: current.rowFieldCols,
    columnFieldCols: current.columnFieldCols,
    filterFieldCols: current.filterFieldCols,
    filterSelectedKeys: current.filterSelectedKeys,
    valueFields: current.valueFields,
    valueFieldsOnRows: current.valueFieldsOnRows === true ? true : undefined,
    destination: {
      kind: "existingSheet",
      startRow: layoutTop,
      startCol: current.destinationCol,
    },
    ...(unconfigured !== undefined ? { unconfiguredPlaceholder: unconfigured } : {}),
  });

  const snapRows = Math.max(current.outputRowCount, out.rowCount);
  const snapCols = Math.max(current.outputColCount, out.colCount);
  pivotSheet.removePivotTableDefinitionById(pivotDefId);
  writePivotResult(pivotSheet, layoutTop, current.destinationCol, out);
  clearPivotOutputOverflow(
    pivotSheet,
    layoutTop,
    current.destinationCol,
    snapRows,
    snapCols,
    out.rowCount,
    out.colCount,
  );
  const nextDef: WorksheetPivotTableDefinition = {
    ...clonePivotDef(current),
    outputRowCount: out.rowCount,
    outputColCount: out.colCount,
  };
  pivotSheet.registerPivotTableDefinition(nextDef);
  return true;
}

export function findPivotTableDefinitionAtCell(
  sheet: Worksheet,
  row: number,
  col: number,
): WorksheetPivotTableDefinition | null {
  for (const p of sheet.getPivotTableDefinitionsSnapshot()) {
    const r0 = pivotLayoutStartRow(p);
    const c0 = p.destinationCol;
    const r1 = r0 + p.outputRowCount - 1;
    const c1 = c0 + p.outputColCount - 1;
    if (row >= r0 && row <= r1 && col >= c0 && col <= c1) {
      return p;
    }
  }
  return null;
}

export class CreatePivotTableCommand implements ICommand {
  readonly id = "feature.createPivotTable";
  readonly label = "创建数据透视表";

  private readonly output: PivotRenderOutput;
  private readonly sourceRange: SelectionRange;
  private readonly pivotDefId: string;
  private createdSheet: Worksheet | null = null;
  private previousActiveSheetIndex = 0;
  private targetSheet: Worksheet | null = null;
  private targetStartRow = 0;
  private targetStartCol = 0;
  private beforeSnapshots: readonly CellSnapshot[] | null = null;

  constructor(
    private readonly workbook: Workbook,
    private readonly sourceSheet: Worksheet,
    private readonly options: PivotTableBuildOptions,
  ) {
    this.sourceRange = normalizeSelectionRange(options.sourceRange);
    this.output = buildPivotRender(sourceSheet, options);
    this.pivotDefId = newPivotDefId();
  }

  execute(): void {
    this.previousActiveSheetIndex = this.workbook.activeSheetIndex;
    if (this.options.destination.kind === "newSheet") {
      this.executeToNewSheet();
      return;
    }
    this.executeToExistingSheet();
  }

  undo(): void {
    if (this.options.destination.kind === "newSheet") {
      this.undoNewSheet();
      return;
    }
    this.undoExistingSheet();
  }

  private executeToNewSheet(): void {
    if (this.options.destination.kind !== "newSheet") {
      return;
    }
    if (this.createdSheet === null) {
      const fallbackName = nextDefaultPivotSheetName(this.workbook);
      const name = this.options.destination.preferredName?.trim() || fallbackName;
      this.createdSheet = new Worksheet(
        name,
        Math.max(200, this.output.rowCount + 24),
        Math.max(32, this.output.colCount + 8),
      );
    }
    if (findSheetIndex(this.workbook, this.createdSheet) < 0) {
      const sourceIdx = findSheetIndex(this.workbook, this.sourceSheet);
      const insertAt = sourceIdx >= 0 ? sourceIdx + 1 : this.workbook.sheetCount;
      this.workbook.insertSheetAt(insertAt, this.createdSheet);
    }
    this.createdSheet.clearPivotTableDefinitions();
    writePivotResult(this.createdSheet, 0, 0, this.output);
    this.registerPivotMetadata(this.createdSheet, 0, 0);
    const idx = findSheetIndex(this.workbook, this.createdSheet);
    if (idx >= 0) {
      this.workbook.activeSheetIndex = idx;
    }
  }

  private undoNewSheet(): void {
    if (this.createdSheet === null) {
      return;
    }
    this.createdSheet.clearPivotTableDefinitions();
    const idx = findSheetIndex(this.workbook, this.createdSheet);
    if (idx >= 0) {
      this.workbook.removeSheetAt(idx);
    }
    this.workbook.activeSheetIndex = this.previousActiveSheetIndex;
  }

  private executeToExistingSheet(): void {
    if (this.options.destination.kind !== "existingSheet") {
      return;
    }
    this.targetSheet = this.options.destination.targetSheet ?? this.sourceSheet;
    this.targetStartRow = this.options.destination.startRow;
    this.targetStartCol = this.options.destination.startCol;
    if (this.beforeSnapshots === null) {
      this.beforeSnapshots = captureSnapshots(
        this.targetSheet,
        this.targetStartRow,
        this.targetStartCol,
        this.output.rowCount,
        this.output.colCount,
      );
    }
    this.targetSheet.removePivotTableDefinitionById(this.pivotDefId);
    writePivotResult(this.targetSheet, this.targetStartRow, this.targetStartCol, this.output);
    this.registerPivotMetadata(this.targetSheet, this.targetStartRow, this.targetStartCol);
  }

  private undoExistingSheet(): void {
    if (this.targetSheet === null || this.beforeSnapshots === null) {
      return;
    }
    this.targetSheet.removePivotTableDefinitionById(this.pivotDefId);
    restoreSnapshots(this.targetSheet, this.beforeSnapshots);
  }

  getSourceRange(): SelectionRange {
    return this.sourceRange;
  }

  getPivotDefinitionId(): string {
    return this.pivotDefId;
  }

  private registerPivotMetadata(dest: Worksheet, destRow: number, destCol: number): void {
    const sourceIdx = findSheetIndex(this.workbook, this.sourceSheet);
    if (sourceIdx < 0) {
      return;
    }
    const filterFieldCols = [...this.options.filterFieldCols];
    const columnFieldCols = [...this.options.columnFieldCols];
    const rowFieldCols = [...this.options.rowFieldCols];
    const valueFieldsForDef = this.options.valueFields.map((v) => ({ ...v }));
    dedupePivotColumnAcrossZones(filterFieldCols, columnFieldCols, rowFieldCols, valueFieldsForDef);
    const def: WorksheetPivotTableDefinition = {
      id: this.pivotDefId,
      name: nextPivotTableExcelName(this.workbook),
      sourceSheetIndex: sourceIdx,
      sourceRange: { ...this.sourceRange },
      hasHeaders: this.options.hasHeaders,
      rowFieldCols,
      columnFieldCols,
      filterFieldCols,
      filterSelectedKeys: normalizeFilterSelectedKeys(
        filterFieldCols,
        this.options.filterSelectedKeys,
      ).map((k) => [...k]),
      valueFields: valueFieldsForDef,
      valueFieldsOnRows:
        this.options.valueFields.length > 1 && this.options.columnFieldCols.length === 0
          ? this.options.valueFieldsOnRows === true
          : undefined,
      destinationRow: destRow,
      destinationCol: destCol,
      outputRowCount: this.output.rowCount,
      outputColCount: this.output.colCount,
    };
    dest.registerPivotTableDefinition(def);
  }
}

type UpdatePivotCmdState = "idle" | "done" | "undone";

/** 在已放置的透视输出区域内更新布局并同步元数据（可撤销）。 */
export class UpdatePivotTableLayoutCommand implements ICommand {
  readonly id = "feature.updatePivotTableLayout";
  readonly label = "更新数据透视表";

  private state: UpdatePivotCmdState = "idle";
  private beforeSnapshots: readonly CellSnapshot[] | null = null;
  private prevDef: WorksheetPivotTableDefinition | null = null;
  private snapRows = 0;
  private snapCols = 0;
  private pivotDestRow = 0;
  private pivotDestCol = 0;
  private nextOutput: PivotRenderOutput | null = null;
  private nextDef: WorksheetPivotTableDefinition | null = null;

  constructor(
    private readonly workbook: Workbook,
    private readonly pivotSheet: Worksheet,
    private readonly pivotDefId: string,
    private readonly layout: {
      readonly sourceRange: SelectionRange;
      readonly hasHeaders: boolean;
      readonly rowFieldCols: readonly number[];
      readonly columnFieldCols: readonly number[];
      readonly filterFieldCols: readonly number[];
      readonly filterSelectedKeys?: readonly (readonly string[])[];
      readonly valueFields: readonly PivotValueFieldSpec[];
      readonly valueFieldsOnRows?: boolean;
    },
  ) {}

  execute(): void {
    const current = this.pivotSheet
      .getPivotTableDefinitionsSnapshot()
      .find((d) => d.id === this.pivotDefId);
    if (current === undefined) {
      return;
    }
    const sourceSheet = this.workbook.getSheet(current.sourceSheetIndex);
    if (sourceSheet === undefined) {
      return;
    }
    if (this.state === "undone") {
      if (this.nextOutput !== null && this.nextDef !== null) {
        this.pivotSheet.removePivotTableDefinitionById(this.pivotDefId);
        writePivotResult(this.pivotSheet, this.pivotDestRow, this.pivotDestCol, this.nextOutput);
        clearPivotOutputOverflow(
          this.pivotSheet,
          this.pivotDestRow,
          this.pivotDestCol,
          this.snapRows,
          this.snapCols,
          this.nextOutput.rowCount,
          this.nextOutput.colCount,
        );
        this.pivotSheet.registerPivotTableDefinition(this.nextDef);
      }
      this.state = "done";
      return;
    }
    if (this.state === "done") {
      return;
    }

    const layoutFilterCols = [...this.layout.filterFieldCols];
    const layoutColumnCols = [...this.layout.columnFieldCols];
    const layoutRowCols = [...this.layout.rowFieldCols];
    const layoutValueFields = this.layout.valueFields.map((v) => ({ ...v }));
    dedupePivotColumnAcrossZones(layoutFilterCols, layoutColumnCols, layoutRowCols, layoutValueFields);
    const mergedFilterKeys = mergePivotFilterSelections(
      current.filterFieldCols,
      current.filterSelectedKeys ?? [],
      layoutFilterCols,
      this.layout.filterSelectedKeys,
    );
    const mergedValueFieldsOnRows =
      layoutColumnCols.length > 0 || layoutValueFields.length <= 1
        ? false
        : this.layout.valueFieldsOnRows !== undefined
          ? this.layout.valueFieldsOnRows
          : (current.valueFieldsOnRows === true);
    const layoutTop = pivotLayoutStartRow(current);
    const buildOpts: PivotTableBuildOptions = {
      sourceRange: this.layout.sourceRange,
      hasHeaders: this.layout.hasHeaders,
      rowFieldCols: layoutRowCols,
      columnFieldCols: layoutColumnCols,
      filterFieldCols: layoutFilterCols,
      filterSelectedKeys: normalizeFilterSelectedKeys(layoutFilterCols, mergedFilterKeys),
      valueFields: layoutValueFields,
      valueFieldsOnRows: mergedValueFieldsOnRows ? true : undefined,
      destination: {
        kind: "existingSheet",
        startRow: layoutTop,
        startCol: current.destinationCol,
      },
    };
    const out = buildPivotRender(sourceSheet, buildOpts);
    this.snapRows = Math.max(current.outputRowCount, out.rowCount);
    this.snapCols = Math.max(current.outputColCount, out.colCount);
    this.pivotDestRow = layoutTop;
    this.pivotDestCol = current.destinationCol;
    this.beforeSnapshots = captureSnapshots(
      this.pivotSheet,
      this.pivotDestRow,
      this.pivotDestCol,
      this.snapRows,
      this.snapCols,
    );
    this.prevDef = clonePivotDef(current);

    this.pivotSheet.removePivotTableDefinitionById(this.pivotDefId);
    writePivotResult(this.pivotSheet, this.pivotDestRow, this.pivotDestCol, out);
    clearPivotOutputOverflow(
      this.pivotSheet,
      this.pivotDestRow,
      this.pivotDestCol,
      this.snapRows,
      this.snapCols,
      out.rowCount,
      out.colCount,
    );

    const persistedFilterKeys = normalizeFilterSelectedKeys(layoutFilterCols, mergedFilterKeys);
    const nextDef: WorksheetPivotTableDefinition = {
      id: current.id,
      name: current.name,
      sourceSheetIndex: current.sourceSheetIndex,
      sourceRange: { ...normalizeSelectionRange(this.layout.sourceRange) },
      hasHeaders: this.layout.hasHeaders,
      rowFieldCols: [...layoutRowCols],
      columnFieldCols: [...layoutColumnCols],
      filterFieldCols: [...layoutFilterCols],
      filterSelectedKeys: persistedFilterKeys.map((k) => [...k]),
      valueFields: layoutValueFields.map((v) => ({ ...v })),
      valueFieldsOnRows:
        layoutValueFields.length > 1 && layoutColumnCols.length === 0
          ? mergedValueFieldsOnRows === true
          : undefined,
      destinationRow: current.destinationRow,
      destinationCol: current.destinationCol,
      outputRowCount: out.rowCount,
      outputColCount: out.colCount,
      pageFilterStartRow: current.pageFilterStartRow,
    };
    this.pivotSheet.registerPivotTableDefinition(nextDef);
    this.nextOutput = out;
    this.nextDef = nextDef;
    this.state = "done";
  }

  undo(): void {
    if (this.state !== "done" || this.beforeSnapshots === null || this.prevDef === null) {
      return;
    }
    this.pivotSheet.removePivotTableDefinitionById(this.pivotDefId);
    restoreSnapshots(this.pivotSheet, this.beforeSnapshots);
    this.pivotSheet.registerPivotTableDefinition(this.prevDef);
    this.state = "undone";
  }
}

type PivotFilterCmdState = "idle" | "done" | "undone";

/** 仅更新透视表筛选项并重新汇总（可撤销）。 */
export class UpdatePivotTableFiltersCommand implements ICommand {
  readonly id = "feature.updatePivotTableFilters";
  readonly label = "更新数据透视表筛选";

  private state: PivotFilterCmdState = "idle";
  private beforeSnapshots: readonly CellSnapshot[] | null = null;
  private prevDef: WorksheetPivotTableDefinition | null = null;
  private snapRows = 0;
  private snapCols = 0;
  private pivotDestRow = 0;
  private pivotDestCol = 0;
  private nextOutput: PivotRenderOutput | null = null;
  private nextDef: WorksheetPivotTableDefinition | null = null;

  constructor(
    private readonly workbook: Workbook,
    private readonly pivotSheet: Worksheet,
    private readonly pivotDefId: string,
    private readonly nextFilterSelectedKeys: readonly (readonly string[])[],
  ) {}

  execute(): void {
    const current = this.pivotSheet
      .getPivotTableDefinitionsSnapshot()
      .find((d) => d.id === this.pivotDefId);
    if (current === undefined) {
      return;
    }
    const sourceSheet = this.workbook.getSheet(current.sourceSheetIndex);
    if (sourceSheet === undefined) {
      return;
    }
    if (this.state === "undone") {
      if (this.nextOutput !== null && this.nextDef !== null) {
        this.pivotSheet.removePivotTableDefinitionById(this.pivotDefId);
        writePivotResult(this.pivotSheet, this.pivotDestRow, this.pivotDestCol, this.nextOutput);
        clearPivotOutputOverflow(
          this.pivotSheet,
          this.pivotDestRow,
          this.pivotDestCol,
          this.snapRows,
          this.snapCols,
          this.nextOutput.rowCount,
          this.nextOutput.colCount,
        );
        this.pivotSheet.registerPivotTableDefinition(this.nextDef);
      }
      this.state = "done";
      return;
    }
    if (this.state === "done") {
      return;
    }

    const normalized = normalizeFilterSelectedKeys(current.filterFieldCols, this.nextFilterSelectedKeys);
    const layoutTop = pivotLayoutStartRow(current);
    const buildOpts: PivotTableBuildOptions = {
      sourceRange: current.sourceRange,
      hasHeaders: current.hasHeaders,
      rowFieldCols: current.rowFieldCols,
      columnFieldCols: current.columnFieldCols,
      filterFieldCols: current.filterFieldCols,
      filterSelectedKeys: normalized,
      valueFields: current.valueFields,
      valueFieldsOnRows: current.valueFieldsOnRows === true ? true : undefined,
      destination: {
        kind: "existingSheet",
        startRow: layoutTop,
        startCol: current.destinationCol,
      },
    };
    const out = buildPivotRender(sourceSheet, buildOpts);
    this.snapRows = Math.max(current.outputRowCount, out.rowCount);
    this.snapCols = Math.max(current.outputColCount, out.colCount);
    this.pivotDestRow = layoutTop;
    this.pivotDestCol = current.destinationCol;
    this.beforeSnapshots = captureSnapshots(
      this.pivotSheet,
      this.pivotDestRow,
      this.pivotDestCol,
      this.snapRows,
      this.snapCols,
    );
    this.prevDef = clonePivotDef(current);

    this.pivotSheet.removePivotTableDefinitionById(this.pivotDefId);
    writePivotResult(this.pivotSheet, this.pivotDestRow, this.pivotDestCol, out);
    clearPivotOutputOverflow(
      this.pivotSheet,
      this.pivotDestRow,
      this.pivotDestCol,
      this.snapRows,
      this.snapCols,
      out.rowCount,
      out.colCount,
    );

    const nextDef: WorksheetPivotTableDefinition = {
      id: current.id,
      name: current.name,
      sourceSheetIndex: current.sourceSheetIndex,
      sourceRange: { ...current.sourceRange },
      hasHeaders: current.hasHeaders,
      rowFieldCols: [...current.rowFieldCols],
      columnFieldCols: [...current.columnFieldCols],
      filterFieldCols: [...current.filterFieldCols],
      filterSelectedKeys: normalized.map((k) => [...k]),
      valueFields: current.valueFields.map((v) => ({ ...v })),
      valueFieldsOnRows: current.valueFieldsOnRows,
      destinationRow: current.destinationRow,
      destinationCol: current.destinationCol,
      outputRowCount: out.rowCount,
      outputColCount: out.colCount,
      pageFilterStartRow: current.pageFilterStartRow,
    };
    this.pivotSheet.registerPivotTableDefinition(nextDef);
    this.nextOutput = out;
    this.nextDef = nextDef;
    this.state = "done";
  }

  undo(): void {
    if (this.state !== "done" || this.beforeSnapshots === null || this.prevDef === null) {
      return;
    }
    this.pivotSheet.removePivotTableDefinitionById(this.pivotDefId);
    restoreSnapshots(this.pivotSheet, this.beforeSnapshots);
    this.pivotSheet.registerPivotTableDefinition(this.prevDef);
    this.state = "undone";
  }
}
