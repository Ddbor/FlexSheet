import { PluginBase, PLUGIN_SERVICE_KEYS, type PluginContext } from "@flexsheet/core";
import type { CellEditor } from "@flexsheet/editor";

import type { FlexSheet } from "./flex-sheet.js";
import {
  runClipboardCopy,
  runClipboardCut,
  runClipboardPaste,
} from "./clipboard/clipboard-run.js";
import { isEditableKeydownTarget } from "./keyboard-editable-target.js";

export interface ClipboardPluginOptions {
  readonly canvas: HTMLCanvasElement;
  readonly getFlexSheet: () => FlexSheet;
  /** 快捷键监听目标；通常为 `chromeRoot`，否则为画布。 */
  readonly keyTarget?: HTMLElement;
}

/** 工厂函数，与 `useUndoRedo` 命名风格一致。 */
export function useClipboard(options: ClipboardPluginOptions): ClipboardPlugin {
  return new ClipboardPlugin(options);
}

/**
 * Ctrl+C / Ctrl+X / Ctrl+V（Windows）与 Cmd+C / Cmd+X / Cmd+V（macOS）；
 * 在单元格内联编辑、编辑栏或其它输入控件聚焦时交由浏览器默认行为，不拦截。
 */
export class ClipboardPlugin extends PluginBase {
  readonly name = "flexsheet.clipboard";

  private ctx: PluginContext | null = null;
  private readonly keyTarget: HTMLElement;
  private readonly getFlexSheet: () => FlexSheet;

  constructor(options: ClipboardPluginOptions) {
    super();
    this.keyTarget = options.keyTarget ?? options.canvas;
    this.getFlexSheet = options.getFlexSheet;
  }

  override install(ctx: PluginContext): void {
    this.ctx = ctx;
  }

  override activate(): void {
    this.keyTarget.addEventListener("keydown", this.onKeyDown, true);
  }

  override deactivate(): void {
    this.keyTarget.removeEventListener("keydown", this.onKeyDown, true);
  }

  override destroy(): void {
    this.ctx = null;
  }

  private readonly onKeyDown = (ev: KeyboardEvent): void => {
    if (isEditableKeydownTarget(ev)) {
      return;
    }
    const ctx = this.ctx;
    if (ctx === null) {
      return;
    }
    const editor = ctx.get<CellEditor>(PLUGIN_SERVICE_KEYS.cellEditor);
    if (editor?.isEditing() === true) {
      return;
    }
    const flex = this.getFlexSheet();
    if (ev.key === "Escape") {
      if (flex.clearClipboardMarquee()) {
        ev.preventDefault();
        ev.stopPropagation();
      }
      return;
    }
    const mod = ev.ctrlKey || ev.metaKey;
    if (!mod || ev.altKey || ev.shiftKey) {
      return;
    }
    const key = ev.key.toLowerCase();
    if (key !== "c" && key !== "x" && key !== "v") {
      return;
    }
    const sheet = flex.workbook.getActiveSheet();
    if (sheet === undefined) {
      return;
    }
    ev.preventDefault();
    ev.stopPropagation();
    if (key === "c") {
      void runClipboardCopy(flex, sheet);
      return;
    }
    if (key === "x") {
      void runClipboardCut(flex, sheet);
      return;
    }
    void runClipboardPaste(flex, sheet);
  };
}
