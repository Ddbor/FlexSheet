import type { FlexSheetLike, RibbonCommandEvent } from "./ribbon-types.js";

/**
 * 将部分 Ribbon 命令映射到 FlexSheet（非「视图」选项卡逻辑可放此处）。
 * 「视图」选项卡由 `ViewRibbonController` 统一处理。
 */
export function applyRibbonCommandToFlexSheet(ev: RibbonCommandEvent, fs: FlexSheetLike): boolean {
  switch (ev.id) {
    case "home.undo.back":
      fs.undo();
      return true;
    case "home.undo.forward":
      fs.redo();
      return true;
    default:
      return false;
  }
}
