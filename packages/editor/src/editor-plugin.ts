import { PluginBase, PLUGIN_SERVICE_KEYS, type PluginContext } from "@flexsheet/core";
import { CellEditor, type CellEditorOptions } from "./cell-editor.js";

/**
 * 单元格浮层编辑插件。
 */
export class EditorPlugin extends PluginBase {
  readonly name = "flexsheet.editor";

  private editor: CellEditor | null = null;

  constructor(private readonly options: CellEditorOptions) {
    super();
  }

  getCellEditor(): CellEditor {
    if (this.editor === null) {
      throw new Error("EditorPlugin: editor not initialized");
    }
    return this.editor;
  }

  override install(_ctx: PluginContext): void {
    this.editor = new CellEditor(this.options);
    _ctx.register(PLUGIN_SERVICE_KEYS.cellEditor, this.editor);
  }

  override destroy(): void {
    if (this.editor !== null) {
      this.editor.dispose();
      this.editor = null;
    }
  }
}
