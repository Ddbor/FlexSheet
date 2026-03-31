/**
 * FlexSheet 公共 API — 框架无关，供宿主或 @flexsheet/react、@flexsheet/vue 等适配。
 */

export const FLEXSHEET_VERSION = "0.0.1" as const;

export {
  Cell,
  Workbook,
  Worksheet,
  normalizeSelectionRange,
  selectionRangeContains,
  type CellAddress,
  type CellScalar,
  type CellStyle,
  type SelectionPaintSnapshot,
  type SelectionRange,
  type WorkbookChangeListener,
  type WorksheetChangeListener,
  Workspace,
  CommandManager,
  EventEmitter,
  UIRegistry,
  PluginBase,
  PLUGIN_SERVICE_KEYS,
  type IPlugin,
  type PluginContext,
  type ICommand,
  type ContextMenuItem,
  type ToolbarSlot,
  type EventHandler,
} from "@flexsheet/core";

export {
  bodyPaintExtents,
  buildFrozenLayout,
  CanvasRenderer,
  clampScroll,
  computeScrollLimits,
  RendererPlugin,
  visibleScrollableCellRange,
  type CanvasRendererOptions,
  type FrozenLayout,
  type RendererPluginOptions,
  type ViewportScrollLimits,
  type VisibleScrollRange,
} from "@flexsheet/renderer";

export {
  FlexSheet,
  createDefaultWorkbook,
  type FlexSheetOptions,
} from "./flex-sheet.js";

export { UndoRedoPlugin, useUndoRedo, type UndoRedoPluginOptions } from "./undo-redo-plugin.js";

export {
  createDefaultDarkTheme,
  createDefaultLightTheme,
  ThemePlugin,
  type SheetTheme,
} from "@flexsheet/theme";

export { columnIndexToLabel } from "@flexsheet/shared";

export {
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
  downloadXlsxBlob,
  exportWorkbookToXlsxBlob,
  exportWorkbookToXlsxBytes,
  importXlsxToWorkbook,
} from "@flexsheet/import-export";

export {
  FlexSheetRibbon,
  ViewRibbonController,
  applyRibbonCommandToFlexSheet,
  type FlexSheetRibbonOptions,
  type RibbonCommandEvent,
  type RibbonTabId,
  type FlexSheetLike,
  type HomeTabHandles,
  type ViewTabHandles,
} from "@flexsheet/toolbar";
