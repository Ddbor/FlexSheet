import {
  Workbook,
  Worksheet,
  computeTableFormatCellStyle,
  TABLE_ACCENT_PALETTES,
  normalizeSelectionRange,
  type CellBorderSide,
  type CellStyle,
  type ParsedTableStyleCommand,
  type SelectionRange,
} from "@flexsheet/core";
import { recalcWorksheet } from "@flexsheet/formula";

const thinBorder = (colorArgb = "FF7F7F7F"): CellBorderSide => ({
  kind: "thin",
  colorArgb,
});

/** 将「套用表格格式」写入区域并注册表样式元数据（与 Ribbon 中「表格样式」一致）。 */
function applyTableFormatRegion(
  sheet: Worksheet,
  range: SelectionRange,
  parsed: ParsedTableStyleCommand,
  hasHeaders: boolean,
): void {
  const n = normalizeSelectionRange(range);
  const palette = TABLE_ACCENT_PALETTES[parsed.col]!;
  for (let row = n.startRow; row <= n.endRow; row++) {
    for (let col = n.startCol; col <= n.endCol; col++) {
      if (sheet.isMergeCoveredCell(row, col)) {
        continue;
      }
      const st = computeTableFormatCellStyle(parsed, palette, n, hasHeaders, row, col);
      sheet.setCellStyle(row, col, st);
    }
  }
  sheet.registerTableStyleRegion(n, parsed, hasHeaders);
}

const titleStyle: CellStyle = {
  bold: true,
  fontSizePt: 14,
  hAlign: "center",
  vAlign: "middle",
};

function fillOverview(sheet: Worksheet): void {
  sheet.setCellLiteral(0, 0, "FlexSheet 功能演示 — 底栏切换各工作表查看");
  sheet.applyMergeForSelection(
    { startRow: 0, endRow: 0, startCol: 0, endCol: 7 },
    "mergeCenter",
  );
  sheet.setCellStyle(0, 0, titleStyle);

  const lines = [
    "【样式】请打开「单元格样式」「表格样式」「条件格式」工作表：分别演示单元格样式、套用表格格式、条件格式规则。",
    "【Ribbon】开始（字体/对齐/数字格式/条件格式/套用表格格式）、插入、公式、数据、视图等选项卡。",
    "【编辑栏】名称框、fx、单元格内容与公式输入（与 Ribbon、快捷键、右键菜单联动）。",
    "【网格】行列标、选择、填充柄、合并单元格、行列宽/隐藏、冻结窗格（视图）。",
    "【筛选】「筛选与表格」工作表：列筛选 + 表格样式联动；与 Ribbon「数据」筛选一致。",
    "【数据】排序与筛选、删除重复项、数据验证、模拟分析、合并计算、创建数据透视表等。",
    "【透视】「透视数据源」提供示例数据；打开本页后自动生成「透视表示例」工作表。",
    "【导入导出】文件菜单：FlexSheet JSON、Excel .xlsx（多表、样式、合并、条件格式、透视元数据）。",
    "【底栏】工作表标签、水平滚动、缩放滑块与状态栏。",
  ];
  let r = 2;
  for (const line of lines) {
    sheet.setCellLiteral(r, 0, line);
    r++;
  }
  sheet.setCellLiteral(
    r + 1,
    0,
    "提示：使用 Ctrl+Z / Ctrl+Y 撤销重做；Ctrl+C / Ctrl+V 剪贴板；右键打开上下文菜单。",
  );
}

function fillFormulas(sheet: Worksheet): void {
  sheet.setCellLiteral(0, 0, "说明");
  sheet.setCellLiteral(0, 1, "数值");
  sheet.setCellLiteral(0, 2, "备注");
  sheet.setCellStyle(0, 0, { bold: true });
  sheet.setCellStyle(0, 1, { bold: true });
  sheet.setCellStyle(0, 2, { bold: true });

  sheet.setCellLiteral(1, 0, "一月");
  sheet.setCellLiteral(1, 1, 120);
  sheet.setCellLiteral(2, 0, "二月");
  sheet.setCellLiteral(2, 1, 240);
  sheet.setCellLiteral(3, 0, "三月");
  sheet.setCellLiteral(3, 1, 180);

  sheet.setCellLiteral(4, 0, "小计");
  sheet.setCellFormula(4, 1, "=SUM(B2:B4)");

  sheet.setCellLiteral(5, 0, "达标提示");
  sheet.setCellFormula(5, 1, '=IF(B5>=500,"达标","未达标")');

  recalcWorksheet(sheet);
}

/** 单元格样式：字体、填充、边框、对齐、数字格式、合并（开始 → 字体 / 对齐 / 数字 / 样式库）。 */
function fillCellStyles(sheet: Worksheet): void {
  sheet.setCellLiteral(0, 0, "单元格样式演示（Ribbon「开始」→ 字体 / 对齐 / 数字）");
  sheet.applyMergeForSelection(
    { startRow: 0, endRow: 0, startCol: 0, endCol: 8 },
    "mergeCenter",
  );
  sheet.setCellStyle(0, 0, { ...titleStyle, fontSizePt: 12 });

  sheet.setCellLiteral(2, 0, "一、字体与填充");
  sheet.setCellStyle(2, 0, { bold: true });

  sheet.setCellLiteral(3, 0, "粗体 13pt");
  sheet.setCellStyle(3, 0, { bold: true, fontSizePt: 13 });
  sheet.setCellLiteral(3, 2, "斜体 + 单下划线");
  sheet.setCellStyle(3, 2, { italic: true, underline: "single" });
  sheet.setCellLiteral(3, 5, "删除线");
  sheet.setCellStyle(3, 5, { strikethrough: true });
  sheet.setCellLiteral(3, 7, "字体色 + 填充色");
  sheet.setCellStyle(3, 7, { fgArgb: "FFC00000", fillArgb: "FFFFFFE0" });

  sheet.setCellLiteral(5, 0, "二、边框（外框细线）");
  sheet.setCellStyle(5, 0, { bold: true });
  sheet.setCellLiteral(6, 0, "带边框单元格");
  sheet.setCellStyle(6, 0, {
    borderTop: thinBorder(),
    borderLeft: thinBorder(),
    borderBottom: thinBorder(),
    borderRight: thinBorder(),
  });

  sheet.setCellLiteral(8, 0, "三、对齐、缩进与换行");
  sheet.setCellStyle(8, 0, { bold: true });
  sheet.setCellLiteral(9, 0, "水平居中");
  sheet.setCellStyle(9, 0, { hAlign: "center" });
  sheet.setCellLiteral(9, 2, "右对齐");
  sheet.setCellStyle(9, 2, { hAlign: "right" });
  sheet.setCellLiteral(9, 4, "顶对齐");
  sheet.setCellStyle(9, 4, { vAlign: "top", wrapText: true });
  sheet.setCellLiteral(10, 0, "自动换行：这是一段较长的说明文字，用于演示在列宽固定时折行显示。");
  sheet.setCellStyle(10, 0, { wrapText: true });
  sheet.setCellLiteral(10, 5, "缩进 2");
  sheet.setCellStyle(10, 5, { indentLevel: 2 });

  sheet.setCellLiteral(12, 0, "四、数字格式");
  sheet.setCellStyle(12, 0, { bold: true });
  sheet.setCellLiteral(13, 0, "百分比");
  sheet.setCellLiteral(13, 1, 0.125);
  sheet.setCellStyle(13, 1, { numberFormat: "0.00%" });
  sheet.setCellLiteral(13, 3, "货币");
  sheet.setCellLiteral(13, 4, 1999.5);
  sheet.setCellStyle(13, 4, { numberFormat: "¥#,##0.00" });
  sheet.setCellLiteral(13, 6, "日期");
  sheet.setCellLiteral(13, 7, 44927);
  sheet.setCellStyle(13, 7, { numberFormat: "yyyy/m/d" });
  sheet.setCellLiteral(13, 9, "科学计数");
  sheet.setCellLiteral(13, 10, 12345.678);
  sheet.setCellStyle(13, 10, { numberFormat: "0.00E+00" });

  sheet.setCellLiteral(15, 0, "五、合并单元格");
  sheet.setCellStyle(15, 0, { bold: true });
  sheet.setCellLiteral(16, 0, "纵向合并标题");
  sheet.setCellLiteral(16, 1, "横向合并内容区");
  sheet.applyMergeForSelection(
    { startRow: 16, endRow: 17, startCol: 0, endCol: 0 },
    "mergeCenter",
  );
  sheet.applyMergeForSelection(
    { startRow: 16, endRow: 16, startCol: 1, endCol: 5 },
    "mergeCenter",
  );
  sheet.setCellStyle(16, 1, { hAlign: "center", vAlign: "middle" });
}

/** 表格样式：多套「套用表格格式」浅色 / 中等 / 深色（与 Ribbon「开始 → 套用表格格式」一致）。 */
function fillTableStyles(sheet: Worksheet): void {
  sheet.setCellLiteral(0, 0, "表格样式演示（开始 → 套用表格格式）");
  sheet.applyMergeForSelection(
    { startRow: 0, endRow: 0, startCol: 0, endCol: 14 },
    "mergeCenter",
  );
  sheet.setCellStyle(0, 0, { ...titleStyle, fontSizePt: 12 });

  const mkRows = (): readonly (readonly [string, string, string, number])[] => [
    ["华东", "Q1", "产品A", 1200],
    ["华北", "Q1", "产品B", 980],
    ["华东", "Q2", "产品A", 1350],
    ["华南", "Q2", "产品C", 760],
  ];

  sheet.setCellLiteral(2, 0, "浅色 · 表样式第 1 行第 3 列主题（黄）");
  sheet.setCellStyle(2, 0, { bold: true });
  let r0 = 3;
  const rows = mkRows();
  sheet.setCellLiteral(r0, 0, "区域");
  sheet.setCellLiteral(r0, 1, "季度");
  sheet.setCellLiteral(r0, 2, "SKU");
  sheet.setCellLiteral(r0, 3, "销售额");
  for (let c = 0; c <= 3; c++) {
    sheet.setCellStyle(r0, c, { bold: true });
  }
  let r = r0 + 1;
  for (const row of rows) {
    sheet.setCellLiteral(r, 0, row[0]);
    sheet.setCellLiteral(r, 1, row[1]);
    sheet.setCellLiteral(r, 2, row[2]);
    sheet.setCellLiteral(r, 3, row[3]);
    sheet.setCellStyle(r, 3, { numberFormat: "#,##0" });
    r++;
  }
  applyTableFormatRegion(
    sheet,
    { startRow: r0, startCol: 0, endRow: r - 1, endCol: 3 },
    { section: "light", row: 0, col: 2 },
    true,
  );

  sheet.setCellLiteral(2, 5, "中等深度 · 第 2 行第 4 列主题（银灰）");
  sheet.setCellStyle(2, 5, { bold: true });
  const cOff = 5;
  r0 = 3;
  sheet.setCellLiteral(r0, cOff + 0, "区域");
  sheet.setCellLiteral(r0, cOff + 1, "季度");
  sheet.setCellLiteral(r0, cOff + 2, "SKU");
  sheet.setCellLiteral(r0, cOff + 3, "销售额");
  for (let c = 0; c <= 3; c++) {
    sheet.setCellStyle(r0, cOff + c, { bold: true });
  }
  r = r0 + 1;
  for (const row of rows) {
    sheet.setCellLiteral(r, cOff + 0, row[0]);
    sheet.setCellLiteral(r, cOff + 1, row[1]);
    sheet.setCellLiteral(r, cOff + 2, row[2]);
    sheet.setCellLiteral(r, cOff + 3, row[3]);
    sheet.setCellStyle(r, cOff + 3, { numberFormat: "#,##0" });
    r++;
  }
  applyTableFormatRegion(
    sheet,
    { startRow: r0, startCol: cOff + 0, endRow: r - 1, endCol: cOff + 3 },
    { section: "medium", row: 1, col: 3 },
    true,
  );

  sheet.setCellLiteral(2, 10, "深色 · 第 2 行第 7 列主题（绿）");
  sheet.setCellStyle(2, 10, { bold: true });
  const cOff2 = 10;
  r0 = 3;
  sheet.setCellLiteral(r0, cOff2 + 0, "区域");
  sheet.setCellLiteral(r0, cOff2 + 1, "季度");
  sheet.setCellLiteral(r0, cOff2 + 2, "SKU");
  sheet.setCellLiteral(r0, cOff2 + 3, "销售额");
  for (let c = 0; c <= 3; c++) {
    sheet.setCellStyle(r0, cOff2 + c, { bold: true });
  }
  r = r0 + 1;
  for (const row of rows) {
    sheet.setCellLiteral(r, cOff2 + 0, row[0]);
    sheet.setCellLiteral(r, cOff2 + 1, row[1]);
    sheet.setCellLiteral(r, cOff2 + 2, row[2]);
    sheet.setCellLiteral(r, cOff2 + 3, row[3]);
    sheet.setCellStyle(r, cOff2 + 3, { numberFormat: "#,##0" });
    r++;
  }
  applyTableFormatRegion(
    sheet,
    { startRow: r0, startCol: cOff2 + 0, endRow: r - 1, endCol: cOff2 + 3 },
    { section: "dark", row: 1, col: 6 },
    true,
  );

  sheet.setCellLiteral(11, 0, "说明：套用表格格式会写入边框与条纹并注册表区域，后续排序/筛选会保持表样式。");
}

/** 条件格式：色阶、数据条、突出显示、前 / 后项（开始 → 条件格式）。 */
function fillConditionalFormat(sheet: Worksheet): void {
  sheet.setCellLiteral(0, 0, "条件格式演示（开始 → 条件格式）");
  sheet.applyMergeForSelection(
    { startRow: 0, endRow: 0, startCol: 0, endCol: 10 },
    "mergeCenter",
  );
  sheet.setCellStyle(0, 0, { ...titleStyle, fontSizePt: 12 });

  sheet.setCellLiteral(2, 0, "1）双色刻度");
  sheet.setCellStyle(2, 0, { bold: true });
  sheet.setCellLiteral(3, 0, "分数");
  for (let c = 0; c < 5; c++) {
    sheet.setCellLiteral(4, c, 20 + c * 15);
  }
  sheet.addConditionalFormatRule({
    id: "demo-cf-two-scale",
    range: { startRow: 4, startCol: 0, endRow: 4, endCol: 4 },
    uiFamily: "twoColorScale",
    classicType: "colorScale",
    formatPreset: "none",
    cfTwoColorMin: { type: "lowest", value: "", colorArgb: "FFFFFFFF" },
    cfTwoColorMax: { type: "highest", value: "", colorArgb: "FF4472C4" },
  });

  sheet.setCellLiteral(6, 0, "2）三色刻度");
  sheet.setCellStyle(6, 0, { bold: true });
  for (let c = 0; c < 5; c++) {
    sheet.setCellLiteral(7, c, 10 + c * 20);
  }
  sheet.addConditionalFormatRule({
    id: "demo-cf-three-scale",
    range: { startRow: 7, startCol: 0, endRow: 7, endCol: 4 },
    uiFamily: "threeColorScale",
    classicType: "colorScale",
    formatPreset: "none",
    cfThreeColorMin: { type: "lowest", value: "", colorArgb: "FFF8696B" },
    cfThreeColorMid: { type: "percentile", value: "50", colorArgb: "FFFFEB84" },
    cfThreeColorMax: { type: "highest", value: "", colorArgb: "FF63BE7B" },
  });

  sheet.setCellLiteral(9, 0, "3）数据条");
  sheet.setCellStyle(9, 0, { bold: true });
  for (let c = 0; c < 5; c++) {
    sheet.setCellLiteral(10, c, 5 + c * 7);
  }
  sheet.addConditionalFormatRule({
    id: "demo-cf-databar",
    range: { startRow: 10, startCol: 0, endRow: 10, endCol: 4 },
    uiFamily: "dataBar",
    classicType: "dataBar",
    formatPreset: "none",
    cfDataBarMin: { type: "lowest", value: "" },
    cfDataBarMax: { type: "highest", value: "" },
    cfDataBarFillKind: "gradient",
    cfDataBarPositiveFillArgb: "FF638EC6",
    cfDataBarNegativeFillArgb: "FFFFB3B3",
    cfDataBarBorderKind: "solid",
    cfDataBarPositiveBorderArgb: "FF638EC6",
    cfDataBarNegativeBorderArgb: "FFFF9999",
    cfDataBarAxisPosition: "automatic",
    cfDataBarAxisColorArgb: "FF808080",
  });

  sheet.setCellLiteral(12, 0, "4）突出显示单元格（值大于 5）");
  sheet.setCellStyle(12, 0, { bold: true });
  sheet.setCellLiteral(13, 0, "状态值");
  sheet.setCellLiteral(14, 0, 3);
  sheet.setCellLiteral(14, 1, 8);
  sheet.setCellLiteral(14, 2, 12);
  sheet.addConditionalFormatRule({
    id: "demo-cf-classic-gt",
    range: { startRow: 14, startCol: 0, endRow: 14, endCol: 2 },
    uiFamily: "classic",
    classicType: "cellsThatContain",
    cellsThatContainKind: "cellValue",
    valueOperator: "greaterThan",
    value1: "5",
    formatPreset: "greenFillDarkGreenText",
  });

  sheet.setCellLiteral(16, 0, "5）最前 / 最后规则（前 2 名高亮）");
  sheet.setCellStyle(16, 0, { bold: true });
  sheet.setCellLiteral(17, 0, "排名");
  for (let c = 0; c < 5; c++) {
    sheet.setCellLiteral(18, c, 12 + c * 13);
  }
  sheet.addConditionalFormatRule({
    id: "demo-cf-top2",
    range: { startRow: 18, startCol: 0, endRow: 18, endCol: 4 },
    uiFamily: "classic",
    classicType: "topBottomRanked",
    topBottomKind: "top",
    topBottomN: 2,
    formatPreset: "yellowFillDarkYellowText",
  });
}

function fillFilterAndTable(sheet: Worksheet): void {
  sheet.setCellLiteral(0, 0, "区域");
  sheet.setCellLiteral(0, 1, "负责人");
  sheet.setCellLiteral(0, 2, "状态");
  sheet.setCellLiteral(0, 3, "预算");
  for (let c = 0; c <= 3; c++) {
    sheet.setCellStyle(0, c, { bold: true });
  }

  const rows: readonly (readonly [string, string, string, number])[] = [
    ["华东", "张三", "进行中", 12000],
    ["华北", "李四", "完成", 8500],
    ["华东", "王五", "进行中", 15000],
    ["华南", "赵六", "暂停", 6000],
    ["华北", "钱七", "完成", 9200],
  ];
  let r = 1;
  for (const row of rows) {
    sheet.setCellLiteral(r, 0, row[0]);
    sheet.setCellLiteral(r, 1, row[1]);
    sheet.setCellLiteral(r, 2, row[2]);
    sheet.setCellLiteral(r, 3, row[3]);
    sheet.setCellStyle(r, 3, { numberFormat: "#,##0" });
    r++;
  }

  const parsed: ParsedTableStyleCommand = { section: "light", row: 0, col: 1 };
  const range = normalizeSelectionRange({ startRow: 0, endRow: r - 1, startCol: 0, endCol: 3 });
  applyTableFormatRegion(sheet, range, parsed, true);

  for (let col = 0; col <= 3; col++) {
    sheet.enableColumnAutoFilterFromSelection(1, col, range);
  }
}

function fillPivotSource(sheet: Worksheet): void {
  sheet.setCellLiteral(0, 0, "类别");
  sheet.setCellLiteral(0, 1, "项目");
  sheet.setCellLiteral(0, 2, "金额");
  for (let c = 0; c <= 2; c++) {
    sheet.setCellStyle(0, c, { bold: true });
  }
  const data: readonly (readonly [string, string, number])[] = [
    ["办公", "纸张", 400],
    ["办公", "文具", 220],
    ["差旅", "机票", 3200],
    ["差旅", "酒店", 1800],
    ["市场", "广告", 9000],
    ["市场", "活动", 4500],
  ];
  let r = 1;
  for (const row of data) {
    sheet.setCellLiteral(r, 0, row[0]);
    sheet.setCellLiteral(r, 1, row[1]);
    sheet.setCellLiteral(r, 2, row[2]);
    r++;
  }
}

/**
 * 构建演示用工作簿：多工作表覆盖公式、样式、合并、条件格式、表格与筛选、透视数据源等能力。
 * 透视结果表由 `main.ts` 在挂载后通过 `CreatePivotTableCommand` 生成。
 */
export function createDemoWorkbook(): Workbook {
  const wb = new Workbook();

  const overview = new Worksheet("功能导览", 40, 12);
  wb.addSheet(overview);
  fillOverview(overview);

  const formulas = new Worksheet("公式示例", 60, 16);
  wb.addSheet(formulas);
  fillFormulas(formulas);

  const cellStyles = new Worksheet("单元格样式", 56, 14);
  wb.addSheet(cellStyles);
  fillCellStyles(cellStyles);

  const tableStyles = new Worksheet("表格样式", 24, 18);
  wb.addSheet(tableStyles);
  fillTableStyles(tableStyles);

  const cf = new Worksheet("条件格式", 48, 16);
  wb.addSheet(cf);
  fillConditionalFormat(cf);

  const table = new Worksheet("筛选与表格", 60, 16);
  wb.addSheet(table);
  fillFilterAndTable(table);

  const pivotSrc = new Worksheet("透视数据源", 80, 16);
  wb.addSheet(pivotSrc);
  fillPivotSource(pivotSrc);

  wb.activeSheetIndex = 0;
  return wb;
}
