import type { CellScalar, CellStyle } from "./cell.js";
import { pivotLayoutStartRow, type WorksheetPivotTableDefinition } from "./pivot-table-model.js";
import type { Worksheet } from "./worksheet.js";

/** 与 Excel 中文界面一致的未配置透视提示（字段窗格为空时）。 */
const PIVOT_UNCONFIGURED_HINT = "在区域内单击可以使用数据透视表";

export function isUnconfiguredPivotDefinition(def: WorksheetPivotTableDefinition): boolean {
  return (
    def.rowFieldCols.length === 0 &&
    def.columnFieldCols.length === 0 &&
    def.filterFieldCols.length === 0 &&
    def.valueFields.length === 0
  );
}

/**
 * 生成「尚未放置任何字段」的透视占位格矩阵（与导入后 / 刷新时写入单元格一致）。
 */
export function buildUnconfiguredPivotPlaceholderMatrix(
  rowCount: number,
  colCount: number,
  title: string,
): {
  readonly rowCount: number;
  readonly colCount: number;
  readonly values: readonly (readonly CellScalar[])[];
  readonly styles: readonly (readonly (CellStyle | null)[])[];
} {
  const rows = Math.max(1, rowCount);
  const cols = Math.max(1, colCount);
  const values: CellScalar[][] = Array.from({ length: rows }, () =>
    Array.from<CellScalar>({ length: cols }).fill(null),
  );
  const styles: (CellStyle | null)[][] = Array.from({ length: rows }, () =>
    Array.from<CellStyle | null>({ length: cols }).fill(null),
  );

  const headerStyle: CellStyle = {
    bold: true,
    fillArgb: "FFBDD7EE",
    hAlign: "left",
    vAlign: "middle",
    fontSizePt: 12,
  };
  const instructStyle: CellStyle = {
    fillArgb: "FFFFFFFF",
    wrapText: true,
    hAlign: "left",
    vAlign: "middle",
    fontSizePt: 11,
  };
  const bodyStyleLeft: CellStyle = {
    fillArgb: "FFF2F2F2",
    hAlign: "left",
    vAlign: "middle",
  };
  const bodyStyleInner: CellStyle = {
    fillArgb: "FFFFFFFF",
    hAlign: "left",
    vAlign: "middle",
  };

  values[0]![0] = title;
  for (let c = 0; c < cols; c++) {
    styles[0]![c] = { ...headerStyle };
  }

  if (rows >= 2) {
    values[1]![0] = PIVOT_UNCONFIGURED_HINT;
    for (let c = 0; c < cols; c++) {
      styles[1]![c] = { ...instructStyle };
    }
  }

  const startBody = rows >= 2 ? 2 : 1;
  for (let r = startBody; r < rows; r++) {
    for (let c = 0; c < cols; c++) {
      values[r]![c] = null;
      styles[r]![c] = c === 0 ? { ...bodyStyleLeft } : { ...bodyStyleInner };
    }
  }

  return { rowCount: rows, colCount: cols, values, styles };
}

/** 将未配置透视的占位内容写入工作表输出区域（会扩展网格）。 */
export function writeUnconfiguredPivotPlaceholderToSheet(
  sheet: Worksheet,
  def: WorksheetPivotTableDefinition,
): void {
  if (!isUnconfiguredPivotDefinition(def)) {
    return;
  }
  const layoutTop = pivotLayoutStartRow(def);
  const c0 = def.destinationCol;
  const out = buildUnconfiguredPivotPlaceholderMatrix(
    def.outputRowCount,
    def.outputColCount,
    def.name,
  );
  const needRows = layoutTop + out.rowCount;
  const needCols = c0 + out.colCount;
  if (needRows > sheet.rowCount || needCols > sheet.colCount) {
    sheet.setGridSize(Math.max(sheet.rowCount, needRows), Math.max(sheet.colCount, needCols));
  }
  sheet.batch(() => {
    for (let r = 0; r < out.rowCount; r++) {
      for (let c = 0; c < out.colCount; c++) {
        const rr = layoutTop + r;
        const cc = c0 + c;
        sheet.setCellLiteral(rr, cc, out.values[r]?.[c] ?? null);
        const st = out.styles[r]?.[c];
        sheet.setCellStyle(rr, cc, st !== undefined && st !== null ? { ...st } : null);
      }
    }
  });
}
