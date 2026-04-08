/**
 * UI 扩展点：工具栏插槽、右键菜单项、Ribbon 占位（由宿主或插件填充）。
 */

/** 内置右键菜单图标键（由宿主将对应 SVG 画在菜单项左侧）。 */
export type ContextMenuBuiltinIconId = "cut" | "copy" | "paste";

export interface ContextMenuItem {
  readonly id: string;
  readonly label: string;
  readonly order?: number;
  readonly disabled?: boolean;
  /** 可选：与 `@flexsheet/toolbar` 中剪贴板图标一致的内置图标。 */
  readonly icon?: ContextMenuBuiltinIconId;
  readonly onSelect?: () => void;
}

/** 右键菜单中的水平分割线（按 `order` 与菜单项一起排序）。 */
export interface ContextMenuSeparator {
  readonly kind: "separator";
  readonly id: string;
  readonly order?: number;
}

/** 菜单项或分割线，供宿主与插件组合注册。 */
export type ContextMenuEntry = ContextMenuItem | ContextMenuSeparator;

export interface ToolbarSlot {
  readonly id: string;
  readonly element: HTMLElement;
}

export class UIRegistry {
  private readonly toolbarSlots = new Map<string, ToolbarSlot>();
  private readonly contextMenuByScope = new Map<string, ContextMenuEntry[]>();

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
  registerContextMenuItems(scope: string, items: readonly ContextMenuEntry[]): () => void {
    this.contextMenuByScope.set(scope, [...items]);
    return () => {
      this.contextMenuByScope.delete(scope);
    };
  }

  getContextMenuItems(scope: string): ContextMenuEntry[] {
    return this.contextMenuByScope.get(scope)?.slice() ?? [];
  }
}
