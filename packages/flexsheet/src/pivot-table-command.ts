import {
  normalizeSelectionRange,
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
  readonly valueFields: readonly PivotValueFieldSpec[];
  readonly destination: PivotTableDestination;
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

function aggregateLabel(kind: PivotAggregateKind): string {
  switch (kind) {
    case "sum":
      return "求和";
    case "count":
      return "计数";
    case "average":
      return "平均值";
    case "max":
      return "最大值";
    case "min":
      return "最小值";
    default:
      return "值";
  }
}

export function dataFieldCaption(agg: PivotAggregateKind, valueFieldName: string): string {
  return `${aggregateLabel(agg)}项:${valueFieldName}`;
}

function styleForValueAggregate(agg: PivotAggregateKind): CellStyle | null {
  if (agg === "count") {
    return { numberFormat: "#,##0" };
  }
  return { numberFormat: "#,##0.00" };
}

function buildPivotRender(sheet: Worksheet, options: PivotTableBuildOptions): PivotRenderOutput {
  const n = normalizeSelectionRange(options.sourceRange);
  const fields = collectFields(sheet, n, options.hasHeaders);
  const fieldNameByCol = new Map<number, string>(fields.map((it) => [it.col, it.name]));

  const rowCols = [...options.rowFieldCols];
  const colDimCols = [...options.columnFieldCols];
  let valueFields = [...options.valueFields];
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

  const matrix = new Map<string, Map<string, PivotAccumulator[]>>();
  const rowKeySet = new Set<string>();
  const colKeySet = new Set<string>();

  const rStart = options.hasHeaders ? n.startRow + 1 : n.startRow;
  const rEnd = n.endRow;
  for (let r = rStart; r <= rEnd; r++) {
    const rowKey = rowCompositeKey(sheet, r, rowCols);
    const colKey = colDimCols.length === 0 ? "__single__" : colCompositeKey(sheet, r, colDimCols);
    rowKeySet.add(rowKey);
    colKeySet.add(colKey);
    let rowMap = matrix.get(rowKey);
    if (rowMap === undefined) {
      rowMap = new Map<string, PivotAccumulator[]>();
      matrix.set(rowKey, rowMap);
    }
    let accs = rowMap.get(colKey);
    if (accs === undefined) {
      accs = valueFields.map(() => createAccumulator());
      rowMap.set(colKey, accs);
    }
    for (let vi = 0; vi < valueFields.length; vi++) {
      const vf = valueFields[vi]!;
      pushAggregate(accs[vi]!, sheet.getCell(r, vf.col).value);
    }
  }

  const rowKeys = [...rowKeySet].sort((a, b) => a.localeCompare(b, "zh-Hans-CN"));
  const colKeys = [...colKeySet].sort((a, b) => a.localeCompare(b, "zh-Hans-CN"));

  const showColDim = colDimCols.length > 0;
  const singleValueNoColDim = !showColDim && valueFields.length === 1;
  const multiValueNoColDim = !showColDim && valueFields.length > 1;

  const colKeysOrdered = colKeys.length === 0 ? ["__single__"] : colKeys;
  const dataColCount = showColDim ? colKeysOrdered.length : valueFields.length;
  const rowTotalColCount = singleValueNoColDim ? 1 : 0;
  const tableCols = 1 + dataColCount + rowTotalColCount;
  const tableRows = 1 + rowKeys.length + 1;

  const values: CellScalar[][] = Array.from({ length: tableRows }, () =>
    Array.from<CellScalar>({ length: tableCols }).fill(null),
  );
  const styles: (CellStyle | null)[][] = Array.from({ length: tableRows }, () =>
    Array.from<CellStyle | null>({ length: tableCols }).fill(null),
  );

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
  const colDimFirstName =
    colDimCols.length > 0 ? (fieldNameByCol.get(colDimCols[0]!) ?? "列") : "";

  values[0][0] =
    showColDim && colDimCols.length > 1
      ? `${rowHeaderLabel}\\${colDimCols.map((c) => fieldNameByCol.get(c) ?? "").join("\\")}`
      : showColDim
        ? `${rowHeaderLabel}\\${colDimFirstName}`
        : rowHeaderLabel;
  styles[0][0] = headStyle;

  if (showColDim) {
    const vf0 = valueFields[0]!;
    const vfName = fieldNameByCol.get(vf0.col) ?? "值";
    for (let ci = 0; ci < colKeysOrdered.length; ci++) {
      const ck = colKeysOrdered[ci]!;
      values[0][ci + 1] = ck === "__single__" ? vfName : ck;
      styles[0][ci + 1] = headStyle;
    }
    values[0][tableCols - 1] = "总计";
    styles[0][tableCols - 1] = headStyle;
  } else if (multiValueNoColDim) {
    for (let vi = 0; vi < valueFields.length; vi++) {
      const vf = valueFields[vi]!;
      const nm = fieldNameByCol.get(vf.col) ?? "值";
      values[0][vi + 1] = dataFieldCaption(vf.aggregate, nm);
      styles[0][vi + 1] = headStyle;
    }
  } else {
    const vf0 = valueFields[0]!;
    const vfName = fieldNameByCol.get(vf0.col) ?? "值";
    values[0][1] = dataFieldCaption(vf0.aggregate, vfName);
    styles[0][1] = headStyle;
    values[0][2] = "总计";
    styles[0][2] = headStyle;
  }

  const colTotals: PivotAccumulator[] = colKeysOrdered.map(() => createAccumulator());
  const allTotal: PivotAccumulator[] = valueFields.map(() => createAccumulator());

  for (let ri = 0; ri < rowKeys.length; ri++) {
    const rowKey = rowKeys[ri]!;
    const rowMap = matrix.get(rowKey) ?? new Map<string, PivotAccumulator[]>();
    values[ri + 1][0] = rowKey.split(KEY_SEP).join(" ");
    styles[ri + 1][0] = null;
    const rowTotalsAcc: PivotAccumulator[] = valueFields.map(() => createAccumulator());

    for (let ci = 0; ci < colKeysOrdered.length; ci++) {
      const colKey = colKeysOrdered[ci]!;
      const accs = rowMap.get(colKey);
      if (showColDim) {
        const vf0 = valueFields[0]!;
        const acc = accs?.[0];
        const out = acc === undefined ? null : reduceAggregate(vf0.aggregate, acc);
        values[ri + 1][ci + 1] = out;
        styles[ri + 1][ci + 1] = styleForValueAggregate(vf0.aggregate);
        if (acc !== undefined) {
          mergeAccumulator(rowTotalsAcc[0]!, acc);
          mergeAccumulator(colTotals[ci]!, acc);
          mergeAccumulator(allTotal[0]!, acc);
        }
      } else if (multiValueNoColDim) {
        for (let vi = 0; vi < valueFields.length; vi++) {
          const vf = valueFields[vi]!;
          const acc = accs?.[vi];
          const out = acc === undefined ? null : reduceAggregate(vf.aggregate, acc);
          values[ri + 1][vi + 1] = out;
          styles[ri + 1][vi + 1] = styleForValueAggregate(vf.aggregate);
          if (acc !== undefined) {
            mergeAccumulator(rowTotalsAcc[vi]!, acc);
            mergeAccumulator(allTotal[vi]!, acc);
          }
        }
      } else {
        const vf0 = valueFields[0]!;
        const acc = accs?.[0];
        const out = acc === undefined ? null : reduceAggregate(vf0.aggregate, acc);
        values[ri + 1][1] = out;
        styles[ri + 1][1] = styleForValueAggregate(vf0.aggregate);
        if (acc !== undefined) {
          mergeAccumulator(rowTotalsAcc[0]!, acc);
          mergeAccumulator(colTotals[ci]!, acc);
          mergeAccumulator(allTotal[0]!, acc);
        }
        values[ri + 1][2] = reduceAggregate(vf0.aggregate, rowTotalsAcc[0]!);
        styles[ri + 1][2] = totalStyle;
      }
    }

    if (showColDim) {
      const vf0 = valueFields[0]!;
      values[ri + 1][tableCols - 1] = reduceAggregate(vf0.aggregate, rowTotalsAcc[0]!);
      styles[ri + 1][tableCols - 1] = totalStyle;
    }
  }

  values[tableRows - 1][0] = "总计";
  styles[tableRows - 1][0] = totalStyle;
  if (showColDim) {
    const vf0 = valueFields[0]!;
    for (let ci = 0; ci < colKeysOrdered.length; ci++) {
      const colTotal = colTotals[ci]!;
      values[tableRows - 1][ci + 1] = reduceAggregate(vf0.aggregate, colTotal);
      styles[tableRows - 1][ci + 1] = totalStyle;
    }
    values[tableRows - 1][tableCols - 1] = reduceAggregate(vf0.aggregate, allTotal[0]!);
    styles[tableRows - 1][tableCols - 1] = totalStyle;
  } else if (multiValueNoColDim) {
    for (let vi = 0; vi < valueFields.length; vi++) {
      const vf = valueFields[vi]!;
      values[tableRows - 1][vi + 1] = reduceAggregate(vf.aggregate, allTotal[vi]!);
      styles[tableRows - 1][vi + 1] = totalStyle;
    }
  } else {
    const vf0 = valueFields[0]!;
    values[tableRows - 1][1] = reduceAggregate(vf0.aggregate, allTotal[0]!);
    styles[tableRows - 1][1] = totalStyle;
    values[tableRows - 1][2] = reduceAggregate(vf0.aggregate, allTotal[0]!);
    styles[tableRows - 1][2] = totalStyle;
  }

  return {
    rowCount: tableRows,
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
    valueFields: def.valueFields.map((v) => ({ ...v })),
  };
}

export function findPivotTableDefinitionAtCell(
  sheet: Worksheet,
  row: number,
  col: number,
): WorksheetPivotTableDefinition | null {
  for (const p of sheet.getPivotTableDefinitionsSnapshot()) {
    const r0 = p.destinationRow;
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
      this.workbook.addSheet(this.createdSheet);
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
    this.targetSheet = this.sourceSheet;
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
    const def: WorksheetPivotTableDefinition = {
      id: this.pivotDefId,
      name: nextPivotTableExcelName(this.workbook),
      sourceSheetIndex: sourceIdx,
      sourceRange: { ...this.sourceRange },
      hasHeaders: this.options.hasHeaders,
      rowFieldCols: [...this.options.rowFieldCols],
      columnFieldCols: [...this.options.columnFieldCols],
      filterFieldCols: [...this.options.filterFieldCols],
      valueFields: this.options.valueFields.map((v) => ({ ...v })),
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
      readonly valueFields: readonly PivotValueFieldSpec[];
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
        writePivotResult(
          this.pivotSheet,
          this.pivotDestRow,
          this.pivotDestCol,
          this.nextOutput,
        );
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

    const buildOpts: PivotTableBuildOptions = {
      sourceRange: this.layout.sourceRange,
      hasHeaders: this.layout.hasHeaders,
      rowFieldCols: this.layout.rowFieldCols,
      columnFieldCols: this.layout.columnFieldCols,
      filterFieldCols: this.layout.filterFieldCols,
      valueFields: this.layout.valueFields,
      destination: {
        kind: "existingSheet",
        startRow: current.destinationRow,
        startCol: current.destinationCol,
      },
    };
    const out = buildPivotRender(sourceSheet, buildOpts);
    this.snapRows = Math.max(current.outputRowCount, out.rowCount);
    this.snapCols = Math.max(current.outputColCount, out.colCount);
    this.pivotDestRow = current.destinationRow;
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
      sourceRange: { ...normalizeSelectionRange(this.layout.sourceRange) },
      hasHeaders: this.layout.hasHeaders,
      rowFieldCols: [...this.layout.rowFieldCols],
      columnFieldCols: [...this.layout.columnFieldCols],
      filterFieldCols: [...this.layout.filterFieldCols],
      valueFields: this.layout.valueFields.map((v) => ({ ...v })),
      destinationRow: current.destinationRow,
      destinationCol: current.destinationCol,
      outputRowCount: out.rowCount,
      outputColCount: out.colCount,
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
