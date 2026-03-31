/**
 * 命令栈（Undo / Redo），供业务与插件统一接入。
 */

export interface ICommand {
  readonly id?: string;
  /** 用于 Ribbon 历史下拉等 UI 展示 */
  readonly label?: string;
  execute(): void;
  undo(): void;
}

export class CommandManager {
  private readonly undoStack: ICommand[] = [];
  private readonly redoStack: ICommand[] = [];
  private readonly changeListeners = new Set<() => void>();
  private maxDepth = 500;

  setMaxHistoryDepth(depth: number): void {
    this.maxDepth = Math.max(1, depth);
    this.trimStacks();
  }

  /** 撤销/重做栈或深度变化时通知（用于工具栏按钮禁用态同步）。 */
  subscribe(listener: () => void): () => void {
    this.changeListeners.add(listener);
    return () => {
      this.changeListeners.delete(listener);
    };
  }

  /** 从底到顶：最早可撤销项在前。 */
  getUndoLabels(): readonly string[] {
    return this.undoStack.map((c) => c.label ?? c.id ?? "操作");
  }

  /** 从下一次重做先到后。 */
  getRedoLabels(): readonly string[] {
    return this.redoStack.map((c) => c.label ?? c.id ?? "操作");
  }

  execute(command: ICommand): void {
    command.execute();
    this.undoStack.push(command);
    if (this.undoStack.length > this.maxDepth) {
      this.undoStack.shift();
    }
    this.redoStack.length = 0;
    this.notifyChange();
  }

  undo(): boolean {
    const cmd = this.undoStack.pop();
    if (cmd === undefined) {
      return false;
    }
    cmd.undo();
    this.redoStack.push(cmd);
    this.notifyChange();
    return true;
  }

  redo(): boolean {
    const cmd = this.redoStack.pop();
    if (cmd === undefined) {
      return false;
    }
    cmd.execute();
    this.undoStack.push(cmd);
    this.notifyChange();
    return true;
  }

  canUndo(): boolean {
    return this.undoStack.length > 0;
  }

  canRedo(): boolean {
    return this.redoStack.length > 0;
  }

  clear(): void {
    this.undoStack.length = 0;
    this.redoStack.length = 0;
    this.notifyChange();
  }

  private notifyChange(): void {
    for (const fn of this.changeListeners) {
      fn();
    }
  }

  private trimStacks(): void {
    while (this.undoStack.length > this.maxDepth) {
      this.undoStack.shift();
    }
  }
}
