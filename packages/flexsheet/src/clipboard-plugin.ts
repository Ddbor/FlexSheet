import {
  PluginBase,
  PLUGIN_SERVICE_KEYS,
  type PluginContext,
  type Worksheet,
} from "@flexsheet/core";
import type { CellEditor } from "@flexsheet/editor";

import type { FlexSheet } from "./flex-sheet.js";
import { CutClearRegionCommand, PasteRegionCommand } from "./clipboard/clipboard-commands.js";
import { matchInternalStyles, setInternalClipboardPayload } from "./clipboard/internal-buffer.js";
import { serializeSelection } from "./clipboard/serialize-selection.js";
import { parseTsv } from "./clipboard/tsv-io.js";

async function writeClipboardText(text: string): Promise<void> {
  try {
    if (typeof navigator !== "undefined" && navigator.clipboard?.writeText !== undefined) {
      await navigator.clipboard.writeText(text);
      return;
    }
  } catch {
    /* 降级 */
  }
  try {
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.style.position = "fixed";
    ta.style.left = "-9999px";
    ta.setAttribute("readonly", "readonly");
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    document.body.removeChild(ta);
  } catch {
    /* 忽略 */
  }
}

async function readClipboardText(): Promise<string | null> {
  try {
    if (typeof navigator !== "undefined" && navigator.clipboard?.readText !== undefined) {
      return await navigator.clipboard.readText();
    }
  } catch {
    return null;
  }
  return null;
}

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
 * 在单元格内联编辑时交由浏览器默认行为，不拦截。
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
      void this.doCopy(sheet);
      return;
    }
    if (key === "x") {
      void this.doCut(flex, sheet);
      return;
    }
    void this.doPaste(flex, sheet);
  };

  private async doCopy(sheet: Worksheet): Promise<void> {
    const flex = this.getFlexSheet();
    const range = flex.selection.getNormalizedRange();
    flex.setClipboardMarquee(range);
    const { tsv, styles } = serializeSelection(sheet, range);
    setInternalClipboardPayload({ tsv, styles });
    await writeClipboardText(tsv);
  }

  private async doCut(flex: FlexSheet, sheet: Worksheet): Promise<void> {
    await this.doCopy(sheet);
    const range = flex.selection.getNormalizedRange();
    const cmd = new CutClearRegionCommand(sheet, range);
    flex.workspace.commands.execute(cmd);
    flex.refresh();
  }

  private async doPaste(flex: FlexSheet, sheet: Worksheet): Promise<void> {
    const text = await readClipboardText();
    if (text === null) {
      return;
    }
    const values = parseTsv(text);
    if (values.length === 0) {
      return;
    }
    const styleGrid = matchInternalStyles(text);
    const ac = flex.selection.getActiveCell();
    const cmd = new PasteRegionCommand(sheet, ac.row, ac.col, values, styleGrid);
    flex.workspace.commands.execute(cmd);
    flex.clearClipboardMarquee();
    flex.refresh();
  }
}
