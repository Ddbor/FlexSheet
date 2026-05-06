import { EXCEL_FUNCTION_CATALOG, type ExcelFunctionCatalogRow } from "./excel-function-catalog.generated.js";
import { FORMULA_BUILDER_COMMON_FUNCTION_ORDER } from "./formula-builder-common-order.js";
import { EXCEL_MS_CATEGORY_BY_NAME } from "./excel-ms-category-map.generated.js";
import type { FormulaBuilderFunctionEntry } from "./formula-builder-panel.js";

/**
 * 其余函数按索引页文档分类（见 `EXCEL_MS_CATEGORY_BY_NAME`）。
 * @see https://support.microsoft.com/zh-cn/office/excel-函数-按字母顺序-b3944572-255d-4efb-bb96-c6d90033e188
 */
const COMMON_SET = new Set(FORMULA_BUILDER_COMMON_FUNCTION_ORDER);

function catalogRowToEntry(row: ExcelFunctionCatalogRow): FormulaBuilderFunctionEntry {
  const common = COMMON_SET.has(row.name);
  const docCategory = EXCEL_MS_CATEGORY_BY_NAME[row.name] ?? "其他";
  return {
    id: row.name,
    name: row.name,
    category: common ? "common" : docCategory,
    description: row.description,
    syntax: row.syntax,
    parameters: row.parameters,
    variadic: row.variadic,
    supportPath: row.supportPath,
  };
}

/** 供公式生成器使用的完整函数列表（说明来自详情页，分组来自索引页分类映射）。 */
export function buildFormulaBuilderFunctionEntries(): FormulaBuilderFunctionEntry[] {
  return EXCEL_FUNCTION_CATALOG.map(catalogRowToEntry);
}
