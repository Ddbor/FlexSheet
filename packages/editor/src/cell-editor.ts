import type { CellScalar } from "@flexsheet/core";
import type { SheetTheme } from "@flexsheet/theme";
import { caretOffsetFromClientX } from "./caret-measure.js";
import { parseEditString } from "./cell-edit-format.js";

export interface CellRectCanvas {
  readonly x: number;
  readonly y: number;
  readonly width: number;
  readonly height: number;
}

export interface CellEditorOptions {
  readonly host: HTMLElement;
  readonly getCanvas: () => HTMLCanvasElement;
  readonly getTheme: () => SheetTheme;
  readonly getCellRect: (row: number, col: number) => CellRectCanvas | null;
  /** 与 Canvas 单元格正文绘制一致（含缩放、加粗），用于对齐与双击测字宽。 */
  readonly getCellFontCss: (row: number, col: number) => string;
  readonly onCommit: (row: number, col: number, value: CellScalar) => void;
  /** 提交或取消编辑后调用（用于将焦点还给表格等）。 */
  readonly onEditEnd?: () => void;
}

export interface BeginEditOptions {
  /** 指针屏幕 X：传入时按该 X 定位插入点（通常由画布双击传入）。 */
  readonly cursorClientX?: number;
  /**
   * 为 false 时插入点在文本末尾、不全选。
   * 未传 `cursorClientX` 时默认 true：全选（便于程序化打开后整格替换）。
   */
  readonly selectAll?: boolean;
}

/**
 * 单个浮层 textarea，绝对定位于宿主内，与 Canvas 单元格对齐。
 */
export class CellEditor {
  private readonly host: HTMLElement;
  private readonly getCanvas: () => HTMLCanvasElement;
  private readonly getTheme: () => SheetTheme;
  private readonly getCellRect: (row: number, col: number) => CellRectCanvas | null;
  private readonly getCellFontCss: (row: number, col: number) => string;
  private readonly onCommit: (row: number, col: number, value: CellScalar) => void;
  private readonly onEditEnd?: () => void;

  private readonly input: HTMLTextAreaElement;
  private readonly measureCanvas: HTMLCanvasElement;

  private editing = false;
  private editingRow = 0;
  private editingCol = 0;
  private skipBlurCommit = false;

  constructor(options: CellEditorOptions) {
    this.host = options.host;
    this.getCanvas = options.getCanvas;
    this.getTheme = options.getTheme;
    this.getCellRect = options.getCellRect;
    this.getCellFontCss = options.getCellFontCss;
    this.onCommit = options.onCommit;
    this.onEditEnd = options.onEditEnd;

    const cs = getComputedStyle(this.host);
    if (cs.position === "static") {
      this.host.style.position = "relative";
    }

    this.input = document.createElement("textarea");
    this.input.autocomplete = "off";
    this.input.spellcheck = false;
    this.input.style.position = "absolute";
    this.input.style.zIndex = "20";
    this.input.style.boxSizing = "border-box";
    this.input.style.margin = "0";
    this.input.style.paddingLeft = "2px";
    this.input.style.paddingRight = "2px";
    this.input.style.paddingTop = "0";
    this.input.style.paddingBottom = "0";
    this.input.style.font = "13px system-ui, -apple-system, sans-serif";
    this.input.style.borderRadius = "0";
    this.input.style.whiteSpace = "pre";
    this.input.style.overflow = "hidden";
    this.input.style.resize = "none";
    this.input.style.display = "none";
    this.applyTheme(this.getTheme());
    this.measureCanvas = document.createElement("canvas");

    this.host.appendChild(this.input);
    this.input.addEventListener("keydown", this.onInputKeyDown);
    this.input.addEventListener("blur", this.onInputBlur);
    this.input.addEventListener("dblclick", this.onInputDblClick);
    this.input.addEventListener("input", this.onInputValueChanged);
  }

  isEditing(): boolean {
    return this.editing;
  }

  applyTheme(theme: SheetTheme): void {
    this.input.style.backgroundColor = theme.editorBg;
    this.input.style.color = theme.editorColor;
    this.input.style.border = `2px solid ${theme.editorBorder}`;
    this.input.style.outline = "none";
  }

  /** 进入编辑；若已在编辑则先提交上一格。 */
  beginEdit(
    row: number,
    col: number,
    initialText: string,
    rect: CellRectCanvas,
    options?: BeginEditOptions,
  ): void {
    if (this.editing) {
      this.commitFromEditor();
    }
    this.editing = true;
    this.editingRow = row;
    this.editingCol = col;
    this.input.value = initialText;
    this.applyRect(rect);
    this.applyFont();
    this.input.style.display = "block";
    this.input.focus();

    if (options?.cursorClientX !== undefined) {
      const cx = options.cursorClientX;
      requestAnimationFrame(() => {
        this.applyTextCaretFromClientX(cx);
      });
    } else if (options?.selectAll === false) {
      const len = initialText.length;
      this.input.setSelectionRange(len, len);
    } else {
      this.input.select();
    }
  }

  /** 布局变化（滚动、resize）时同步 input 位置。 */
  syncLayout(): void {
    if (!this.editing) {
      return;
    }
    const rect = this.getCellRect(this.editingRow, this.editingCol);
    if (rect === null) {
      return;
    }
    this.applyRect(rect);
    this.applyFont();
  }

  /** 外部强制结束编辑并提交（如销毁）。 */
  dispose(): void {
    this.input.removeEventListener("keydown", this.onInputKeyDown);
    this.input.removeEventListener("blur", this.onInputBlur);
    this.input.removeEventListener("dblclick", this.onInputDblClick);
    this.input.removeEventListener("input", this.onInputValueChanged);
    if (this.editing) {
      this.commitFromEditor();
    }
    this.input.remove();
  }

  cancelWithoutCommit(): void {
    if (!this.editing) {
      return;
    }
    this.skipBlurCommit = true;
    this.editing = false;
    this.input.style.display = "none";
    this.input.blur();
    queueMicrotask(() => {
      this.skipBlurCommit = false;
    });
    this.onEditEnd?.();
  }

  /**
   * 宿主 `position: relative` 时，绝对定位子元素相对 **padding 盒** 左上角；
   * 不能简单用 border box 的 getBoundingClientRect 差值，否则有 padding/border 时会偏。
   */
  private hostPaddingBoxOrigin(): { x: number; y: number } {
    const r = this.host.getBoundingClientRect();
    const s = getComputedStyle(this.host);
    const bl = parseFloat(s.borderLeftWidth) || 0;
    const bt = parseFloat(s.borderTopWidth) || 0;
    const pl = parseFloat(s.paddingLeft) || 0;
    const pt = parseFloat(s.paddingTop) || 0;
    return { x: r.left + bl + pl, y: r.top + bt + pt };
  }

  private applyFont(): void {
    this.input.style.font = this.getCellFontCss(this.editingRow, this.editingCol);
  }

  private applyRect(rect: CellRectCanvas): void {
    const canvas = this.getCanvas();
    const cr = canvas.getBoundingClientRect();
    const ho = this.hostPaddingBoxOrigin();
    const left = cr.left - ho.x + rect.x;
    const top = cr.top - ho.y + rect.y;
    this.input.style.left = `${left}px`;
    this.input.style.top = `${top}px`;
    this.input.style.width = `${rect.width}px`;
    this.input.style.height = `${rect.height}px`;
    this.input.style.minWidth = `${rect.width}px`;
    this.input.style.minHeight = `${rect.height}px`;
    const borderTotal = 4;
    const innerH = Math.max(0, rect.height - borderTotal);
    this.input.style.lineHeight = `${Math.max(16, innerH)}px`;
    this.applyAutoWidthByContent();
  }

  private getTextStartScreenX(): number {
    const r = this.input.getBoundingClientRect();
    const s = getComputedStyle(this.input);
    const bl = parseFloat(s.borderLeftWidth) || 0;
    const pl = parseFloat(s.paddingLeft) || 0;
    return r.left + bl + pl;
  }

  private applyTextCaretFromClientX(clientX: number): void {
    if (!this.editing) {
      return;
    }
    const font = this.getCellFontCss(this.editingRow, this.editingCol);
    const originX = this.getTextStartScreenX();
    const offset = caretOffsetFromClientX(this.input.value, clientX, originX, font);
    this.input.setSelectionRange(offset, offset);
  }

  private readonly onInputDblClick = (ev: MouseEvent): void => {
    if (!this.editing) {
      return;
    }
    ev.preventDefault();
    ev.stopPropagation();
    requestAnimationFrame(() => {
      this.applyTextCaretFromClientX(ev.clientX);
    });
  };

  private readonly onInputValueChanged = (): void => {
    this.applyAutoWidthByContent();
  };

  private readonly onInputKeyDown = (ev: KeyboardEvent): void => {
    if (ev.key === "Enter" && !ev.altKey) {
      if (ev.isComposing) {
        return;
      }
      ev.preventDefault();
      ev.stopPropagation();
      this.skipBlurCommit = true;
      this.commitFromEditor();
      this.input.blur();
      queueMicrotask(() => {
        this.skipBlurCommit = false;
      });
      return;
    }
    if (ev.key === "Escape") {
      if (ev.isComposing) {
        return;
      }
      ev.preventDefault();
      ev.stopPropagation();
      this.cancelWithoutCommit();
    }
  };

  private readonly onInputBlur = (): void => {
    if (this.skipBlurCommit) {
      return;
    }
    if (!this.editing) {
      return;
    }
    this.commitFromEditor();
  };

  private commitFromEditor(): void {
    if (!this.editing) {
      return;
    }
    const row = this.editingRow;
    const col = this.editingCol;
    const value = parseEditString(this.input.value);
    this.editing = false;
    this.input.style.display = "none";
    this.onCommit(row, col, value);
    this.onEditEnd?.();
  }

  private applyAutoWidthByContent(): void {
    if (!this.editing) {
      return;
    }
    const rect = this.getCellRect(this.editingRow, this.editingCol);
    if (rect === null) {
      return;
    }
    const ctx = this.measureCanvas.getContext("2d");
    if (ctx === null) {
      return;
    }
    const font = this.getCellFontCss(this.editingRow, this.editingCol);
    ctx.font = font;
    const lines = this.input.value.split("\n");
    let maxLineW = 0;
    for (const line of lines) {
      maxLineW = Math.max(maxLineW, ctx.measureText(line).width);
    }
    const pad = 10;
    const desired = Math.ceil(maxLineW + pad);
    this.input.style.width = `${Math.max(rect.width, desired)}px`;
  }
}
