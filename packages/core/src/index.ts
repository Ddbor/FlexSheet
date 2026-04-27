export {
  Cell,
  applyCellStylePatch,
  type CellAddress,
  type CellBorderKind,
  type CellBorderLinePattern,
  type CellBorderSide,
  type CellHorizontalAlign,
  type CellScalar,
  type CellStyle,
  type CellStylePatch,
  type CellTextOrientation,
  CELL_FILL_PATTERN_TYPES,
  isCellFillPatternType,
  type CellFillPatternType,
  type CellVerticalAlign,
} from "./cell.js";
export { resolveCellBorderStroke, type ResolvedCellBorderStroke } from "./border-line-pattern.js";
export {
  adjustDecimalPlacesInFormat,
  applyCommaStyleFromFormat,
  excelSerialToUtcDate,
  formatCellDisplayWithStyle,
  formatNumberWithExcelCode,
} from "./excel-number-format.js";
export {
  AUTO_FILTER_BLANK_KEY,
  cellToAutoFilterDisplayKey,
  compareAutoFilterDisplayKeys,
} from "./column-auto-filter-keys.js";
export { Workbook, type WorkbookChangeListener } from "./workbook.js";
export {
  Worksheet,
  type ColumnAutoFilterSortHint,
  type ColumnAutoFilterUiKind,
  type ColumnAutoFilterUndoSnapshot,
  type WorksheetChangeListener,
  type WorksheetCustomSortLevel,
} from "./worksheet.js";
export {
  normalizeSelectionRange,
  selectionRangeContains,
  selectionRangesEqualNormalized,
  selectionRangesIntersect,
  type SelectionPaintSnapshot,
  type SelectionRange,
} from "./selection-range.js";
export {
  CF_ICON_GLYPH_PICKER_ORDER,
  CF_ICON_SET_CATALOG,
  cfFormatPresetToOverlay,
  cellMatchesConditionalFormatRule,
  findCfIconSetCatalogEntry,
  resolveConditionalFormattingOverlay,
  resolveDataBarPaintForCell,
  resolveIconSetGlyphForCell,
  resolveThreeColorScaleFillArgb,
  resolveTwoColorScaleFillArgb,
  type CfAverageKind,
  type CfColorScaleEndpoint,
  type CfColorScaleEndpointType,
  type CfCellsThatContainKind,
  type CfDataBarAxisPosition,
  type CfDataBarBorderKind,
  type CfDataBarDirection,
  type CfDataBarFillKind,
  type CfDataBarMaxEndpointType,
  type CfDataBarMinEndpointType,
  type CfDateOccurring,
  type CfFormatPresetId,
  type CfIconGlyphId,
  type CfIconSetCatalogEntry,
  type CfIconSetId,
  type CfIconThresholdRow,
  type CfIconThresholdValueType,
  type CfTextOperator,
  type CfTopBottomKind,
  type CfUniqueKind,
  type CfValueOperator,
  type ConditionalFormattingCellIcon,
  type ConditionalFormattingDataBarPaint,
  type ConditionalFormatRule,
  type ConditionalFormatClassicRuleType,
  type ConditionalFormattingOverlay,
  type ConditionalFormatUiFamily,
} from "./conditional-formatting.js";

export { EventEmitter, type EventHandler } from "./event-emitter.js";
export { CommandManager, type ICommand } from "./command-manager.js";
export {
  UIRegistry,
  type ContextMenuBuiltinIconId,
  type ContextMenuEntry,
  type ContextMenuItem,
  type ContextMenuSeparator,
  type ContextMenuSubItem,
  type ToolbarSlot,
} from "./ui-registry.js";
export { PLUGIN_SERVICE_KEYS, type PluginServiceKey } from "./plugin-keys.js";
export { PluginBase, type IPlugin, type PluginContext } from "./plugin-types.js";
export { Workspace } from "./workspace.js";
export {
  TABLE_ACCENT_PALETTES,
  computeTableFormatCellStyle,
  ooxmlTableStyleNameToParsed,
  parseTableStyleRibbonCommand,
  tableStyleUsesDistinctHeaderRow,
  type ParsedTableStyleCommand,
  type TableAccentPalette,
  type TableStyleSection,
} from "./table-style-presets.js";
export {
  getPivotValueFieldCaption,
  pivotLayoutStartRow,
  type PivotAggregateKind,
  type PivotValueComputed,
  type PivotValueFieldSpec,
  type WorksheetPivotTableDefinition,
} from "./pivot-table-model.js";
export {
  buildUnconfiguredPivotPlaceholderMatrix,
  isUnconfiguredPivotDefinition,
  writeUnconfiguredPivotPlaceholderToSheet,
} from "./pivot-unconfigured-placeholder.js";
