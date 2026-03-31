import { PluginBase, PLUGIN_SERVICE_KEYS, type PluginContext } from "@flexsheet/core";
import type { SelectionModel } from "./selection-model.js";

/**
 * 将 `SelectionModel` 注册到插件上下文，供其他插件协作。
 */
export class SelectionRegistryPlugin extends PluginBase {
  readonly name = "flexsheet.selection";

  constructor(private readonly selection: SelectionModel) {
    super();
  }

  override install(ctx: PluginContext): void {
    ctx.register(PLUGIN_SERVICE_KEYS.selection, this.selection);
  }
}
