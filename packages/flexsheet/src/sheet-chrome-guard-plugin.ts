import { PLUGIN_SERVICE_KEYS, PluginBase, type PluginContext } from "@flexsheet/core";
import type { CellEditor } from "@flexsheet/editor";

export interface SheetChromeGuardPluginOptions {
  /** 包含 Ribbon、编辑栏、表格主体、底部栏等的共同祖先，用于统一拦截浏览器默认行为。 */
  readonly chromeRoot: HTMLElement;
}

export function useSheetChromeGuard(options: SheetChromeGuardPluginOptions): SheetChromeGuardPlugin {
  return new SheetChromeGuardPlugin(options);
}

/**
 * 在表格工作区内禁用浏览器默认右键菜单（输入框等仍保留系统行为），
 * 并拦截常见浏览器快捷键（Ctrl+S 保存网页等），保留与内核约定的 Ctrl+C/X/V/Z/Y/A 等由后续逻辑处理。
 */
export class SheetChromeGuardPlugin extends PluginBase {
  readonly name = "flexsheet.sheetChromeGuard";

  private ctx: PluginContext | null = null;
  private readonly chromeRoot: HTMLElement;

  constructor(options: SheetChromeGuardPluginOptions) {
    super();
    this.chromeRoot = options.chromeRoot;
  }

  override install(ctx: PluginContext): void {
    this.ctx = ctx;
  }

  override activate(): void {
    this.chromeRoot.addEventListener("contextmenu", this.onContextMenuCapture, true);
    this.chromeRoot.addEventListener("keydown", this.onKeyDownCapture, true);
  }

  override deactivate(): void {
    this.chromeRoot.removeEventListener("contextmenu", this.onContextMenuCapture, true);
    this.chromeRoot.removeEventListener("keydown", this.onKeyDownCapture, true);
  }

  override destroy(): void {
    this.ctx = null;
  }

  private readonly onContextMenuCapture = (ev: MouseEvent): void => {
    if (!this.chromeRoot.contains(ev.target as Node)) {
      return;
    }
    if (allowNativeContextMenu(ev)) {
      return;
    }
    ev.preventDefault();
  };

  private readonly onKeyDownCapture = (ev: KeyboardEvent): void => {
    if (!this.chromeRoot.contains(ev.target as Node)) {
      return;
    }
    const ctx = this.ctx;
    const editor = ctx?.get<CellEditor>(PLUGIN_SERVICE_KEYS.cellEditor);
    if (editor?.isEditing() === true) {
      return;
    }
    if (allowNativeKeydownInEditable(ev)) {
      return;
    }
    const mod = ev.ctrlKey || ev.metaKey;
    if (!mod) {
      return;
    }
    if (ev.altKey) {
      return;
    }
    const key = ev.key.toLowerCase();
    const passThrough = new Set(["c", "x", "v", "z", "y", "a"]);
    if (passThrough.has(key)) {
      return;
    }
    ev.preventDefault();
  };
}

function allowNativeContextMenu(ev: MouseEvent): boolean {
  const t = ev.target;
  if (t === null || !(t instanceof Node)) {
    return false;
  }
  const el = t.nodeType === Node.TEXT_NODE ? t.parentElement : (t as HTMLElement);
  if (el === null) {
    return false;
  }
  const host = el.closest("input, textarea, [contenteditable='true']");
  return host !== null;
}

function allowNativeKeydownInEditable(ev: KeyboardEvent): boolean {
  const t = ev.target;
  if (t === null) {
    return false;
  }
  if (t instanceof HTMLInputElement || t instanceof HTMLTextAreaElement) {
    return true;
  }
  if (t instanceof HTMLElement && t.isContentEditable) {
    return true;
  }
  return false;
}
