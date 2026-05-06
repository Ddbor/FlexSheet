import type { ICommand } from "@flexsheet/core";

import { writeClipboardImageFromDataUrl } from "../clipboard/clipboard-io.js";
import { setFloatingPictureClipboard } from "../clipboard/internal-buffer.js";
import type {
  FloatingPicturePastePrepared,
  FloatingPictureSnapshot,
} from "../chrome/floating-picture-layer.js";
import type { FlexSheet } from "../flex-sheet.js";

/** 剪切浮动图：可撤销恢复；复制不进命令栈。 */
export class CutFloatingPictureCommand implements ICommand {
  readonly id = "floatingPicture.cut";
  readonly label = "剪切图片";

  constructor(
    private readonly flex: FlexSheet,
    private readonly snapshot: FloatingPictureSnapshot,
  ) {}

  execute(): void {
    setFloatingPictureClipboard(this.snapshot.dataUrl);
    void writeClipboardImageFromDataUrl(this.snapshot.dataUrl).catch(() => {
      /* 系统剪贴板不可用时仍可用应用内粘贴 */
    });
    this.flex.removeFloatingPictureById(this.snapshot.id);
  }

  undo(): void {
    this.flex.restoreFloatingPictureFromSnapshot(this.snapshot);
  }
}

/** 粘贴浮动图：可撤销删除、重做再插入。 */
export class PasteFloatingPictureCommand implements ICommand {
  readonly id = "floatingPicture.paste";
  readonly label = "粘贴图片";

  private committed: FloatingPictureSnapshot | null = null;

  constructor(
    private readonly flex: FlexSheet,
    private readonly prepared: FloatingPicturePastePrepared,
  ) {}

  execute(): void {
    if (this.committed === null) {
      this.committed = this.flex.insertFloatingPictureFromPrepared(this.prepared);
    } else {
      this.flex.restoreFloatingPictureFromSnapshot(this.committed);
    }
  }

  undo(): void {
    if (this.committed !== null) {
      this.flex.removeFloatingPictureById(this.committed.id);
    }
  }
}
