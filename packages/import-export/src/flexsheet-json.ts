/**
 * FlexSheet 工作簿 JSON：纯数据交换格式，供本软件导出/导入与校验。
 */

import {
  Cell,
  Workbook,
  Worksheet,
  isCellFillPatternType,
  type CellScalar,
  type CellStyle,
} from "@flexsheet/core";
import { recalcWorksheet } from "@flexsheet/formula";

/** 根对象上的格式标识（导入时强校验）。 */
export const FLEXSHEET_JSON_FORMAT = "flexsheet.workbook+json" as const;

/** 与 `FLEXSHEET_JSON_FORMAT` 配套的结构版本。 */
export const FLEXSHEET_JSON_FORMAT_VERSION = 1 as const;

/** `generator.app` 期望值；非此值则视为非本软件导出。 */
export const FLEXSHEET_JSON_GENERATOR_APP = "FlexSheet" as const;

export interface FlexSheetJsonExportOptions {
  readonly includeStyles: boolean;
  readonly includeFormulas: boolean;
  /** 预留：当前数据模型无独立名称表，导出无额外行为。 */
  readonly includeUnusedNames: boolean;
  readonly saveByView: boolean;
  /** 预留：当前无合并元数据。 */
  readonly includeAutoMergedCells: boolean;
  /** 为公式单元格同时导出已缓存的标量值 `v`。 */
  readonly includeCalculationCache: boolean;
  /** 预留：公式仍以 A1 风格存储。 */
  readonly saveFormulasAsR1C1: boolean;
  /** 预留。 */
  readonly includeBoundDataSources: boolean;
  /** 导出仅有样式、无公式且无值的单元格（用于样式占位）。 */
  readonly includeSparseStyledEmpty: boolean;
}

export interface FlexSheetJsonViewState {
  readonly frozenRows: number;
  readonly frozenCols: number;
  readonly scrollX: number;
  readonly scrollY: number;
}

export interface FlexSheetJsonCell {
  readonly r: number;
  readonly c: number;
  readonly v?: CellScalar;
  readonly f?: string;
  readonly s?: CellStyle | null;
}

export interface FlexSheetJsonSheet {
  readonly name: string;
  readonly rowCount: number;
  readonly colCount: number;
  readonly defaultRowHeight?: number;
  readonly defaultColWidth?: number;
  readonly rowHeights?: readonly (readonly [number, number])[];
  readonly colWidths?: readonly (readonly [number, number])[];
  readonly hiddenRows?: readonly number[];
  readonly hiddenCols?: readonly number[];
  readonly cells: readonly FlexSheetJsonCell[];
}

export interface FlexSheetJsonWorkbookPayload {
  readonly activeSheetIndex: number;
  readonly view?: FlexSheetJsonViewState;
  readonly sheets: readonly FlexSheetJsonSheet[];
}

export interface FlexSheetJsonDocument {
  readonly format: typeof FLEXSHEET_JSON_FORMAT;
  readonly formatVersion: typeof FLEXSHEET_JSON_FORMAT_VERSION;
  readonly generator: {
    readonly app: typeof FLEXSHEET_JSON_GENERATOR_APP;
    readonly version?: string;
  };
  readonly workbook: FlexSheetJsonWorkbookPayload;
}

export const DEFAULT_FLEXSHEET_JSON_EXPORT_OPTIONS: FlexSheetJsonExportOptions = {
  includeStyles: true,
  includeFormulas: true,
  includeUnusedNames: true,
  saveByView: false,
  includeAutoMergedCells: false,
  includeCalculationCache: false,
  saveFormulasAsR1C1: false,
  includeBoundDataSources: false,
  includeSparseStyledEmpty: true,
};

/** 打开/导入 JSON 时可选项（与 Backstage「打开」页复选框对应）。 */
export interface FlexSheetJsonImportOptions {
  /** 为 false 时不写入单元格样式。 */
  readonly includeStyles: boolean;
  /** 为 false 时将公式单元格按缓存值 `v` 导入为字面量（无 `v` 则为空）。 */
  readonly includeFormulas: boolean;
  /** 为 false 时不整表重算；公式单元格若 JSON 含 `v` 则用作显示缓存。 */
  readonly recalcAfterImport: boolean;
}

export const DEFAULT_FLEXSHEET_JSON_IMPORT_OPTIONS: FlexSheetJsonImportOptions = {
  includeStyles: true,
  includeFormulas: true,
  recalcAfterImport: true,
};

function isPlainObject(v: unknown): v is Record<string, unknown> {
  return typeof v === "object" && v !== null && !Array.isArray(v);
}

function isCellScalar(v: unknown): v is CellScalar {
  return v === null || typeof v === "string" || typeof v === "number" || typeof v === "boolean";
}

function isCellStyle(v: unknown): v is CellStyle {
  if (!isPlainObject(v)) {
    return false;
  }
  const keys = Object.keys(v);
  const allowed = new Set([
    "bold",
    "italic",
    "fontFamily",
    "fontSizePt",
    "underline",
    "fgArgb",
    "fillArgb",
    "fillPatternType",
    "fillPatternFgArgb",
    "hAlign",
    "vAlign",
    "indentLevel",
    "wrapText",
    "textRotationDegrees",
    "shrinkToFit",
    "textOrientation",
    "numberFormat",
    "locked",
    "formulaHidden",
  ]);
  for (const k of keys) {
    if (!allowed.has(k)) {
      return false;
    }
  }
  if (v.bold !== undefined && typeof v.bold !== "boolean") {
    return false;
  }
  if (v.italic !== undefined && typeof v.italic !== "boolean") {
    return false;
  }
  if (v.fontFamily !== undefined && typeof v.fontFamily !== "string") {
    return false;
  }
  if (
    v.fontSizePt !== undefined &&
    (typeof v.fontSizePt !== "number" || !Number.isFinite(v.fontSizePt) || v.fontSizePt <= 0)
  ) {
    return false;
  }
  if (v.underline !== undefined && v.underline !== "single" && v.underline !== "double") {
    return false;
  }
  if (v.fgArgb !== undefined && typeof v.fgArgb !== "string") {
    return false;
  }
  if (v.fillArgb !== undefined && typeof v.fillArgb !== "string") {
    return false;
  }
  if (v.fillPatternType !== undefined) {
    if (typeof v.fillPatternType !== "string" || !isCellFillPatternType(v.fillPatternType)) {
      return false;
    }
  }
  if (v.fillPatternFgArgb !== undefined) {
    if (typeof v.fillPatternFgArgb !== "string") {
      return false;
    }
    const t = v.fillPatternFgArgb.trim();
    if (t !== "" && !/^[\dA-Fa-f]{8}$/.test(t)) {
      return false;
    }
  }
  if (
    v.hAlign !== undefined &&
    v.hAlign !== "left" &&
    v.hAlign !== "center" &&
    v.hAlign !== "right" &&
    v.hAlign !== "fill" &&
    v.hAlign !== "justify" &&
    v.hAlign !== "distributed" &&
    v.hAlign !== "centerContinuous"
  ) {
    return false;
  }
  if (
    v.vAlign !== undefined &&
    v.vAlign !== "top" &&
    v.vAlign !== "middle" &&
    v.vAlign !== "bottom" &&
    v.vAlign !== "justify" &&
    v.vAlign !== "distributed"
  ) {
    return false;
  }
  if (
    v.indentLevel !== undefined &&
    (typeof v.indentLevel !== "number" ||
      !Number.isFinite(v.indentLevel) ||
      v.indentLevel < 0 ||
      v.indentLevel > 255)
  ) {
    return false;
  }
  if (v.wrapText !== undefined && typeof v.wrapText !== "boolean") {
    return false;
  }
  if (
    v.textRotationDegrees !== undefined &&
    (typeof v.textRotationDegrees !== "number" ||
      !Number.isFinite(v.textRotationDegrees) ||
      v.textRotationDegrees < -90 ||
      v.textRotationDegrees > 90)
  ) {
    return false;
  }
  if (v.shrinkToFit !== undefined && typeof v.shrinkToFit !== "boolean") {
    return false;
  }
  if (v.textOrientation !== undefined) {
    const ok =
      v.textOrientation === "horizontal" ||
      v.textOrientation === "angleUp45" ||
      v.textOrientation === "angleDown45" ||
      v.textOrientation === "verticalStack" ||
      v.textOrientation === "rotateUp90" ||
      v.textOrientation === "rotateDown90";
    if (!ok) {
      return false;
    }
  }
  if (v.numberFormat !== undefined && typeof v.numberFormat !== "string") {
    return false;
  }
  if (v.locked !== undefined && typeof v.locked !== "boolean") {
    return false;
  }
  if (v.formulaHidden !== undefined && typeof v.formulaHidden !== "boolean") {
    return false;
  }
  return true;
}

export type FlexSheetJsonParseResult =
  | { readonly ok: true; readonly doc: FlexSheetJsonDocument }
  | { readonly ok: false; readonly error: string };

export function parseFlexSheetJson(text: string): FlexSheetJsonParseResult {
  let root: unknown;
  try {
    root = JSON.parse(text) as unknown;
  } catch {
    return { ok: false, error: "不是有效的 JSON 文本。" };
  }
  if (!isPlainObject(root)) {
    return { ok: false, error: "根节点必须是 JSON 对象。" };
  }
  if (root.format !== FLEXSHEET_JSON_FORMAT) {
    return { ok: false, error: "缺少或错误的 format 字段，不是 FlexSheet 工作簿 JSON。" };
  }
  if (root.formatVersion !== FLEXSHEET_JSON_FORMAT_VERSION) {
    return {
      ok: false,
      error: `不支持的 formatVersion（仅支持 ${String(FLEXSHEET_JSON_FORMAT_VERSION)}）。`,
    };
  }
  const gen = root.generator;
  if (!isPlainObject(gen) || gen.app !== FLEXSHEET_JSON_GENERATOR_APP) {
    return { ok: false, error: "generator.app 不是 FlexSheet，拒绝导入。" };
  }
  const wbRaw = root.workbook;
  if (!isPlainObject(wbRaw)) {
    return { ok: false, error: "缺少 workbook 对象。" };
  }
  const sheetsRaw = wbRaw.sheets;
  if (!Array.isArray(sheetsRaw) || sheetsRaw.length === 0) {
    return { ok: false, error: "workbook.sheets 必须为非空数组。" };
  }
  const activeSheetIndex = wbRaw.activeSheetIndex;
  if (typeof activeSheetIndex !== "number" || !Number.isInteger(activeSheetIndex)) {
    return { ok: false, error: "workbook.activeSheetIndex 无效。" };
  }

  let view: FlexSheetJsonViewState | undefined;
  if (wbRaw.view !== undefined) {
    if (!isPlainObject(wbRaw.view)) {
      return { ok: false, error: "workbook.view 必须是对象。" };
    }
    const fr = wbRaw.view.frozenRows;
    const fc = wbRaw.view.frozenCols;
    const sx = wbRaw.view.scrollX;
    const sy = wbRaw.view.scrollY;
    if (
      typeof fr !== "number" ||
      typeof fc !== "number" ||
      typeof sx !== "number" ||
      typeof sy !== "number"
    ) {
      return { ok: false, error: "workbook.view 缺少有效的数值字段。" };
    }
    view = {
      frozenRows: Math.max(0, Math.trunc(fr)),
      frozenCols: Math.max(0, Math.trunc(fc)),
      scrollX: sx,
      scrollY: sy,
    };
  }

  const sheets: FlexSheetJsonSheet[] = [];
  for (let si = 0; si < sheetsRaw.length; si++) {
    const sh = sheetsRaw[si];
    if (!isPlainObject(sh)) {
      return { ok: false, error: `工作表 #${si} 不是对象。` };
    }
    const name = sh.name;
    if (typeof name !== "string" || name.trim() === "") {
      return { ok: false, error: `工作表 #${si} 缺少有效名称。` };
    }
    const rowCount = sh.rowCount;
    const colCount = sh.colCount;
    if (typeof rowCount !== "number" || typeof colCount !== "number") {
      return { ok: false, error: `工作表 "${name}" 的 rowCount/colCount 无效。` };
    }
    const r = Math.max(1, Math.trunc(rowCount));
    const c = Math.max(1, Math.trunc(colCount));
    const cellsRaw = sh.cells;
    if (!Array.isArray(cellsRaw)) {
      return { ok: false, error: `工作表 "${name}" 的 cells 必须是数组。` };
    }
    const cells: FlexSheetJsonCell[] = [];
    for (let ci = 0; ci < cellsRaw.length; ci++) {
      const cell = cellsRaw[ci];
      if (!isPlainObject(cell)) {
        return { ok: false, error: `工作表 "${name}" 的 cells[${ci}] 无效。` };
      }
      const rr = cell.r;
      const cc = cell.c;
      if (typeof rr !== "number" || typeof cc !== "number") {
        return { ok: false, error: `工作表 "${name}" 的单元格坐标无效。` };
      }
      const ri = Math.trunc(rr);
      const cj = Math.trunc(cc);
      if (ri < 0 || cj < 0 || ri >= r || cj >= c) {
        return { ok: false, error: `工作表 "${name}" 存在越界单元格 (${ri},${cj})。` };
      }
      const entry: FlexSheetJsonCell = { r: ri, c: cj };
      if ("f" in cell) {
        if (cell.f !== undefined && typeof cell.f !== "string") {
          return { ok: false, error: `工作表 "${name}" 的公式字段类型错误。` };
        }
        if (typeof cell.f === "string" && cell.f.length > 0) {
          Object.assign(entry, { f: cell.f });
        }
      }
      if ("v" in cell && cell.v !== undefined) {
        if (!isCellScalar(cell.v)) {
          return { ok: false, error: `工作表 "${name}" 的单元格值类型错误。` };
        }
        Object.assign(entry, { v: cell.v });
      }
      if ("s" in cell && cell.s !== undefined && cell.s !== null) {
        if (!isCellStyle(cell.s)) {
          return { ok: false, error: `工作表 "${name}" 的单元格样式无效。` };
        }
        Object.assign(entry, { s: { ...cell.s } });
      } else if ("s" in cell && cell.s === null) {
        Object.assign(entry, { s: null });
      }
      cells.push(entry);
    }

    const sheet: FlexSheetJsonSheet = { name: name.trim(), rowCount: r, colCount: c, cells };
    if (typeof sh.defaultRowHeight === "number" && Number.isFinite(sh.defaultRowHeight)) {
      Object.assign(sheet, { defaultRowHeight: Math.max(2, Math.trunc(sh.defaultRowHeight)) });
    }
    if (typeof sh.defaultColWidth === "number" && Number.isFinite(sh.defaultColWidth)) {
      Object.assign(sheet, { defaultColWidth: Math.max(2, Math.trunc(sh.defaultColWidth)) });
    }
    if (Array.isArray(sh.rowHeights)) {
      const pairs: [number, number][] = [];
      for (const p of sh.rowHeights) {
        if (
          Array.isArray(p) &&
          p.length === 2 &&
          typeof p[0] === "number" &&
          typeof p[1] === "number"
        ) {
          pairs.push([Math.trunc(p[0]), Math.max(2, Math.trunc(p[1]))]);
        }
      }
      if (pairs.length > 0) {
        Object.assign(sheet, { rowHeights: pairs });
      }
    }
    if (Array.isArray(sh.colWidths)) {
      const pairs: [number, number][] = [];
      for (const p of sh.colWidths) {
        if (
          Array.isArray(p) &&
          p.length === 2 &&
          typeof p[0] === "number" &&
          typeof p[1] === "number"
        ) {
          pairs.push([Math.trunc(p[0]), Math.max(2, Math.trunc(p[1]))]);
        }
      }
      if (pairs.length > 0) {
        Object.assign(sheet, { colWidths: pairs });
      }
    }
    if (Array.isArray(sh.hiddenRows)) {
      const hr = sh.hiddenRows
        .filter((x): x is number => typeof x === "number")
        .map((x) => Math.trunc(x));
      if (hr.length > 0) {
        Object.assign(sheet, { hiddenRows: hr });
      }
    }
    if (Array.isArray(sh.hiddenCols)) {
      const hc = sh.hiddenCols
        .filter((x): x is number => typeof x === "number")
        .map((x) => Math.trunc(x));
      if (hc.length > 0) {
        Object.assign(sheet, { hiddenCols: hc });
      }
    }
    sheets.push(sheet);
  }

  const clampedActive = Math.max(0, Math.min(Math.trunc(activeSheetIndex), sheets.length - 1));

  const doc: FlexSheetJsonDocument = {
    format: FLEXSHEET_JSON_FORMAT,
    formatVersion: FLEXSHEET_JSON_FORMAT_VERSION,
    generator: {
      app: FLEXSHEET_JSON_GENERATOR_APP,
      ...(typeof gen.version === "string" ? { version: gen.version } : {}),
    },
    workbook: {
      activeSheetIndex: clampedActive,
      ...(view !== undefined ? { view } : {}),
      sheets,
    },
  };
  return { ok: true, doc };
}

export function workbookFromFlexSheetJsonDocument(
  doc: FlexSheetJsonDocument,
  importOptions: FlexSheetJsonImportOptions = DEFAULT_FLEXSHEET_JSON_IMPORT_OPTIONS,
): Workbook {
  const wb = new Workbook();
  for (const js of doc.workbook.sheets) {
    const sh = new Worksheet(js.name, js.rowCount, js.colCount);
    if (js.defaultRowHeight !== undefined) {
      sh.defaultRowHeight = js.defaultRowHeight;
    }
    if (js.defaultColWidth !== undefined) {
      sh.defaultColWidth = js.defaultColWidth;
    }
    sh.batch(() => {
      if (js.rowHeights !== undefined) {
        for (const [row, h] of js.rowHeights) {
          if (row >= 0 && row < sh.rowCount) {
            sh.setRowHeight(row, h);
          }
        }
      }
      if (js.colWidths !== undefined) {
        for (const [col, w] of js.colWidths) {
          if (col >= 0 && col < sh.colCount) {
            sh.setColWidth(col, w);
          }
        }
      }
      if (js.hiddenRows !== undefined) {
        for (const row of js.hiddenRows) {
          if (row >= 0 && row < sh.rowCount) {
            sh.setRowHidden(row, true);
          }
        }
      }
      if (js.hiddenCols !== undefined) {
        for (const col of js.hiddenCols) {
          if (col >= 0 && col < sh.colCount) {
            sh.setColHidden(col, true);
          }
        }
      }
      for (const jc of js.cells) {
        if (importOptions.includeStyles) {
          sh.setCellStyle(jc.r, jc.c, jc.s !== undefined ? jc.s : null);
        } else {
          sh.setCellStyle(jc.r, jc.c, null);
        }
        const hasF = jc.f !== undefined && jc.f.length > 0;
        if (importOptions.includeFormulas && hasF) {
          sh.setCellFormula(jc.r, jc.c, jc.f);
          if (!importOptions.recalcAfterImport && jc.v !== undefined) {
            sh.getCell(jc.r, jc.c).value = jc.v;
          }
        } else if (hasF && !importOptions.includeFormulas) {
          sh.setCellLiteral(jc.r, jc.c, jc.v ?? null);
        } else if (jc.v !== undefined) {
          sh.setCellLiteral(jc.r, jc.c, jc.v);
        }
      }
    });
    wb.addSheet(sh);
  }
  wb.activeSheetIndex = doc.workbook.activeSheetIndex;
  if (importOptions.recalcAfterImport) {
    for (let i = 0; i < wb.sheetCount; i++) {
      const s = wb.getSheet(i);
      if (s !== undefined) {
        recalcWorksheet(s);
      }
    }
  }
  return wb;
}

function shouldExportCell(cell: Cell, opt: FlexSheetJsonExportOptions): boolean {
  const hasFormula = cell.formula !== null && cell.formula.length > 0;
  const hasValue = cell.value !== null && cell.value !== "";
  const hasStyle = cell.style !== null && Object.keys(cell.style).length > 0;
  if (hasFormula && opt.includeFormulas) {
    return true;
  }
  if (hasFormula && !opt.includeFormulas) {
    return hasValue || (opt.includeSparseStyledEmpty && hasStyle && opt.includeStyles);
  }
  if (hasValue) {
    return true;
  }
  if (opt.includeSparseStyledEmpty && hasStyle && opt.includeStyles) {
    return true;
  }
  return false;
}

function cellToJson(cell: Cell, opt: FlexSheetJsonExportOptions): FlexSheetJsonCell {
  const hasFormula = cell.formula !== null && cell.formula.length > 0;
  const r = cell.row;
  const c = cell.col;
  const stylePart =
    opt.includeStyles && cell.style !== null && Object.keys(cell.style).length > 0
      ? ({ s: { ...cell.style } } as const)
      : ({} as const);

  if (hasFormula && opt.includeFormulas) {
    return {
      r,
      c,
      ...stylePart,
      f: cell.formula as string,
      ...(opt.includeCalculationCache ? { v: cell.value } : {}),
    };
  }
  if (hasFormula && !opt.includeFormulas) {
    return { r, c, ...stylePart, v: cell.value };
  }
  return {
    r,
    c,
    ...stylePart,
    ...(cell.value !== null ? { v: cell.value } : {}),
  };
}

export interface FlexSheetJsonSerializeMeta {
  readonly generatorVersion?: string;
  readonly viewState?: FlexSheetJsonViewState;
}

export function serializeWorkbookToJsonDocument(
  workbook: Workbook,
  options: FlexSheetJsonExportOptions,
  meta?: FlexSheetJsonSerializeMeta,
): FlexSheetJsonDocument {
  const sheets: FlexSheetJsonSheet[] = [];
  const list = workbook.getSheets();
  for (const sh of list) {
    const cells: FlexSheetJsonCell[] = [];
    sh.iterateCells((cell) => {
      if (shouldExportCell(cell, options)) {
        cells.push(cellToJson(cell, options));
      }
    });
    cells.sort((a, b) => (a.r !== b.r ? a.r - b.r : a.c - b.c));
    const rowHeights: [number, number][] = [];
    for (let r = 0; r < sh.rowCount; r++) {
      const h = sh.getRowHeight(r);
      if (h !== sh.defaultRowHeight) {
        rowHeights.push([r, h]);
      }
    }
    const colWidths: [number, number][] = [];
    for (let c = 0; c < sh.colCount; c++) {
      const w = sh.getColWidth(c);
      if (w !== sh.defaultColWidth) {
        colWidths.push([c, w]);
      }
    }
    const hiddenRows: number[] = [];
    for (let r = 0; r < sh.rowCount; r++) {
      if (sh.isRowHidden(r)) {
        hiddenRows.push(r);
      }
    }
    const hiddenCols: number[] = [];
    for (let c = 0; c < sh.colCount; c++) {
      if (sh.isColHidden(c)) {
        hiddenCols.push(c);
      }
    }
    const sheetJson: FlexSheetJsonSheet = {
      name: sh.name,
      rowCount: sh.rowCount,
      colCount: sh.colCount,
      defaultRowHeight: sh.defaultRowHeight,
      defaultColWidth: sh.defaultColWidth,
      ...(rowHeights.length > 0 ? { rowHeights } : {}),
      ...(colWidths.length > 0 ? { colWidths } : {}),
      ...(hiddenRows.length > 0 ? { hiddenRows } : {}),
      ...(hiddenCols.length > 0 ? { hiddenCols } : {}),
      cells,
    };
    sheets.push(sheetJson);
  }

  const workbookPayload: FlexSheetJsonWorkbookPayload = {
    activeSheetIndex: Math.max(
      0,
      Math.min(workbook.activeSheetIndex, Math.max(0, sheets.length - 1)),
    ),
    sheets,
    ...(options.saveByView && meta?.viewState !== undefined ? { view: { ...meta.viewState } } : {}),
  };

  return {
    format: FLEXSHEET_JSON_FORMAT,
    formatVersion: FLEXSHEET_JSON_FORMAT_VERSION,
    generator: {
      app: FLEXSHEET_JSON_GENERATOR_APP,
      ...(meta?.generatorVersion !== undefined ? { version: meta.generatorVersion } : {}),
    },
    workbook: workbookPayload,
  };
}

export function downloadJsonText(text: string, filename: string): void {
  if (typeof document === "undefined") {
    return;
  }
  const blob = new Blob([text], { type: "application/json;charset=utf-8" });
  const a = document.createElement("a");
  const url = URL.createObjectURL(blob);
  a.href = url;
  a.download = filename.toLowerCase().endsWith(".json") ? filename : `${filename}.json`;
  a.rel = "noopener";
  a.click();
  URL.revokeObjectURL(url);
}
