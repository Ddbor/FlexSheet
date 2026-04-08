export {
  Cell,
  applyCellStylePatch,
  type CellAddress,
  type CellBorderKind,
  type CellBorderSide,
  type CellHorizontalAlign,
  type CellScalar,
  type CellStyle,
  type CellStylePatch,
  type CellVerticalAlign,
} from "./cell.js";
export { Workbook, type WorkbookChangeListener } from "./workbook.js";
export { Worksheet, type WorksheetChangeListener } from "./worksheet.js";
export {
  normalizeSelectionRange,
  selectionRangeContains,
  type SelectionPaintSnapshot,
  type SelectionRange,
} from "./selection-range.js";

export { EventEmitter, type EventHandler } from "./event-emitter.js";
export { CommandManager, type ICommand } from "./command-manager.js";
export { UIRegistry, type ContextMenuItem, type ToolbarSlot } from "./ui-registry.js";
export { PLUGIN_SERVICE_KEYS, type PluginServiceKey } from "./plugin-keys.js";
export { PluginBase, type IPlugin, type PluginContext } from "./plugin-types.js";
export { Workspace } from "./workspace.js";
