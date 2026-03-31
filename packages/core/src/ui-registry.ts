/**
 * UI 扩展点：工具栏插槽、右键菜单项、Ribbon 占位（由宿主或插件填充）。
 */

export interface ContextMenuItem {
  readonly id: string;
  readonly label: string;
  readonly order?: number;
  readonly disabled?: boolean;
  readonly onSelect?: () => void;
}

export interface ToolbarSlot {
  readonly id: string;
  readonly element: HTMLElement;
}

export class UIRegistry {
  private readonly toolbarSlots = new Map<string, ToolbarSlot>();
  private readonly contextMenuByScope = new Map<string, ContextMenuItem[]>();

  registerToolbarSlot(slot: ToolbarSlot): () => void {
    this.toolbarSlots.set(slot.id, slot);
    return () => {
      this.toolbarSlots.delete(slot.id);
    };
  }

  getToolbarSlot(id: string): ToolbarSlot | undefined {
    return this.toolbarSlots.get(id);
  }

  listToolbarSlotIds(): string[] {
    return [...this.toolbarSlots.keys()];
  }

  /** scope 如 `sheet.body`、`selection` */
  registerContextMenuItems(scope: string, items: readonly ContextMenuItem[]): () => void {
    this.contextMenuByScope.set(scope, [...items]);
    return () => {
      this.contextMenuByScope.delete(scope);
    };
  }

  getContextMenuItems(scope: string): ContextMenuItem[] {
    return this.contextMenuByScope.get(scope)?.slice() ?? [];
  }
}
