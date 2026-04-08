import type { SelectionRange, Worksheet, Workspace } from "@flexsheet/core";

import {
  CutRangeExceptRectCommand,
  PasteRegionCommand,
  getPasteClippedRect,
} from "./clipboard-commands.js";
import { getInternalClipboardPayload, matchInternalStyles, setInternalClipboardPayload } from "./internal-buffer.js";
import { writeClipboardText, readClipboardText } from "./clipboard-io.js";
import { serializeSelection } from "./serialize-selection.js";
import { parseTsv } from "./tsv-io.js";

/** 与 `FlexSheet` 对齐的剪贴板操作目标（避免 `flex-sheet` ↔ 插件循环依赖）。 */
export interface ClipboardRunTarget {
  readonly selection: {
    getNormalizedRange(): SelectionRange;
    getActiveCell(): { readonly row: number; readonly col: number };
  };
  setClipboardMarquee(range: SelectionRange | null): void;
  clearClipboardMarquee(): boolean;
  readonly workspace: Workspace;
  refresh(): void;
  /** Excel 式延迟剪切：粘贴成功且系统剪贴板仍为本应用剪切载荷时再清空源区 */
  setPendingClipboardCut(sheet: Worksheet, range: SelectionRange): void;
  clearPendingClipboardCut(): void;
  getPendingClipboardCut(): { sheet: Worksheet; range: SelectionRange } | null;
}

export async function runClipboardCopy(target: ClipboardRunTarget, sheet: Worksheet): Promise<void> {
  target.clearPendingClipboardCut();
  const range = target.selection.getNormalizedRange();
  target.setClipboardMarquee(range);
  const { tsv, styles } = serializeSelection(sheet, range);
  setInternalClipboardPayload({ tsv, styles });
  await writeClipboardText(tsv);
}

/**
 * 延迟剪切：内容与复制相同仍保留在格内；复制 / Esc / 换表等会取消；
 * 仅在随后粘贴且剪贴板文本仍与内部载荷一致时清空源区（与 Excel 一致）。
 */
export async function runClipboardCut(target: ClipboardRunTarget, sheet: Worksheet): Promise<void> {
  await runClipboardCopy(target, sheet);
  const range = target.selection.getNormalizedRange();
  target.setPendingClipboardCut(sheet, range);
  target.refresh();
}

export async function runClipboardPaste(target: ClipboardRunTarget, sheet: Worksheet): Promise<void> {
  const text = await readClipboardText();
  if (text === null) {
    return;
  }
  const values = parseTsv(text);
  if (values.length === 0) {
    return;
  }
  const styleGrid = matchInternalStyles(text);
  const ac = target.selection.getActiveCell();
  const cmd = new PasteRegionCommand(sheet, ac.row, ac.col, values, styleGrid);

  const pending = target.getPendingClipboardCut();
  const internal = getInternalClipboardPayload();
  const completingMove =
    pending !== null &&
    pending.sheet === sheet &&
    internal !== null &&
    text === internal.tsv;

  target.workspace.commands.execute(cmd);

  if (completingMove && pending !== null) {
    const { h, w } = getPasteClippedRect(sheet, ac.row, ac.col, values);
    const clearCmd = new CutRangeExceptRectCommand(sheet, pending.range, ac.row, ac.col, h, w);
    target.workspace.commands.execute(clearCmd);
  }

  target.clearClipboardMarquee();
  target.refresh();
}
