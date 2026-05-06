/**
 * FlexSheet 公共 API — 框架无关，供宿主或 @flexsheet/react、@flexsheet/vue 等适配。
 */

export const FLEXSHEET_VERSION = "0.0.1" as const;

export {
  Cell,
  Workbook,
  Worksheet,
  applyCellStylePatch,
  normalizeSelectionRange,
  selectionRangeContains,
  selectionRangesEqualNormalized,
  type CellAddress,
  type CellScalar,
  type CellStyle,
  type CellStylePatch,
  type SelectionPaintSnapshot,
  type SelectionRange,
  type WorkbookChangeListener,
  type WorksheetChangeListener,
  type ConditionalFormatRule,
  type ConditionalFormattingOverlay,
  Workspace,
  CommandManager,
  EventEmitter,
  UIRegistry,
  PluginBase,
  PLUGIN_SERVICE_KEYS,
  type IPlugin,
  type PluginContext,
  type ICommand,
  type ContextMenuEntry,
  type ContextMenuItem,
  type ContextMenuSeparator,
  type ContextMenuSubItem,
  type ToolbarSlot,
  type EventHandler,
} from "@flexsheet/core";

export {
  bodyPaintExtents,
  buildFrozenLayout,
  CanvasRenderer,
  clampScroll,
  computeScrollLimits,
  hitTestHeadingPointer,
  RendererPlugin,
  visibleScrollableCellRange,
  type CanvasRendererOptions,
  type FrozenLayout,
  type HeadingHit,
  type HorizontalScrollMetrics,
  type VerticalScrollMetrics,
  type RendererPluginOptions,
  type ViewportScrollLimits,
  type VisibleScrollRange,
} from "@flexsheet/renderer";

export {
  FlexSheet,
  createDefaultWorkbook,
  type FlexSheetOptions,
  type FlexSheetSurfaceHit,
  type SelectionCellDeleteMode,
} from "./flex-sheet.js";

export { mountExcelBottomBar, type MountExcelBottomBarOptions } from "./chrome/excel-bottom-bar.js";

export {
  mountGridVerticalScrollbar,
  type MountGridVerticalScrollbarOptions,
} from "./chrome/grid-vertical-scrollbar.js";

export { UndoRedoPlugin, useUndoRedo, type UndoRedoPluginOptions } from "./plugins/undo-redo-plugin.js";

export {
  CONTEXT_MENU_SCOPE,
  SheetContextMenuPlugin,
  useSheetContextMenu,
  type SheetContextMenuPluginOptions,
} from "./plugins/sheet-context-menu-plugin.js";

export {
  SheetChromeGuardPlugin,
  useSheetChromeGuard,
  type SheetChromeGuardPluginOptions,
} from "./plugins/sheet-chrome-guard-plugin.js";

export { ClipboardPlugin, useClipboard, type ClipboardPluginOptions } from "./plugins/clipboard-plugin.js";

export {
  CutClearRegionCommand,
  CutRangeExceptRectCommand,
  PasteRegionCommand,
  getPasteClippedRect,
} from "./clipboard/clipboard-commands.js";
export {
  DeleteCellsShiftLeftCommand,
  DeleteCellsShiftUpCommand,
  DeleteColsCommand,
  DeleteRowsCommand,
  InsertColsCommand,
  InsertRowsCommand,
  SetColHiddenCommand,
  SetColWidthsInRangeCommand,
  SetColWidthCommand,
  SetRowHeightsInRangeCommand,
  SetRowHeightCommand,
  SetRowHiddenCommand,
} from "./commands/sheet-structure-commands.js";

export {
  ApplyFormatAsTableCommand,
  ApplySelectionCellStylePatchCommand,
  ApplySelectionFontSizeStepCommand,
  ApplySelectionFormatCellsDialogCommand,
  ApplySelectionIndentStepCommand,
} from "./commands/cell-style-commands.js";

export type { FormatCellsBorderState, FormatCellsLineSwatchId } from "./format-cells/format-cells-border.js";

export {
  createDefaultDarkTheme,
  createDefaultLightTheme,
  ThemePlugin,
  type SheetTheme,
} from "@flexsheet/theme";

export { columnIndexToLabel } from "@flexsheet/shared";

export {
  ClearRegionContentsCommand,
  evaluateAst,
  parseFormula,
  recalcWorksheet,
  SetCellValueCommand,
  setCellLiteralAndRecalc,
  setCellValueAndRecalc,
  applyCellSnapshotAndRecalc,
  type CellContentSnapshot,
  stripFormulaEquals,
  toNumber,
  ParseError,
  FormulaEnginePlugin,
  type AstBinary,
  type AstCall,
  type AstNode,
  type AstNumber,
  type AstRange,
  type AstRef,
  type AstUnary,
  type EvalContext,
} from "@flexsheet/formula";

export {
  CellEditor,
  cellScalarToEditString,
  parseEditString,
  EditorPlugin,
  type BeginEditOptions,
  type CellEditorOptions,
  type CellRectCanvas,
} from "@flexsheet/editor";

export { SelectionModel, SelectionRegistryPlugin } from "@flexsheet/selection";

export { ScrollPlugin } from "@flexsheet/scroll";

export { FEATURE_LAYER_PLACEHOLDER } from "./feature/plugin-placeholder.js";

export {
  DEFAULT_FLEXSHEET_JSON_EXPORT_OPTIONS,
  DEFAULT_FLEXSHEET_JSON_IMPORT_OPTIONS,
  downloadJsonText,
  downloadXlsxBlob,
  decodeTextFileBytes,
  DEFAULT_XLSX_EXPORT_OPTIONS,
  exportWorkbookToXlsxBlob,
  exportWorkbookToXlsxBytes,
  FLEXSHEET_JSON_FORMAT,
  FLEXSHEET_JSON_FORMAT_VERSION,
  FLEXSHEET_JSON_GENERATOR_APP,
  importXlsxToWorkbook,
  type XlsxExportOptions,
  type XlsxFloatingPictureExport,
  parseFlexSheetJson,
  serializeWorkbookToJsonDocument,
  workbookFromFlexSheetJsonDocument,
  type FlexSheetJsonDocument,
  type FlexSheetJsonExportOptions,
  type FlexSheetJsonImportOptions,
} from "@flexsheet/import-export";

export {
  FlexSheetRibbon,
  ViewRibbonController,
  applyRibbonCommandToFlexSheet,
  argb8ToCssHex6,
  argb8ToStripeCss,
  cellStyleToRibbonHomeFontChrome,
  cssHexToFillArgb,
  type FlexSheetRibbonOptions,
  type RibbonCommandEvent,
  type RibbonHomeFontChromeState,
  type RibbonTabId,
  type FlexSheetLike,
  type HomeTabHandles,
  type ViewTabHandles,
} from "@flexsheet/toolbar";
