import { PluginBase, PLUGIN_SERVICE_KEYS, type PluginContext } from "@flexsheet/core";
import type { SheetTheme } from "./sheet-theme.js";

/**
 * 将当前主题注册到插件上下文（`PLUGIN_SERVICE_KEYS.theme`），便于其他插件读取。
 */
export class ThemePlugin extends PluginBase {
  readonly name = "flexsheet.theme";

  constructor(private theme: SheetTheme) {
    super();
  }

  override install(ctx: PluginContext): void {
    ctx.register(PLUGIN_SERVICE_KEYS.theme, this.theme);
  }

  setTheme(theme: SheetTheme): void {
    this.theme = theme;
  }
}
