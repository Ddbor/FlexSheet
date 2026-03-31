import { PluginBase, type PluginContext } from "@flexsheet/core";

/**
 * 公式引擎插件占位：在 `install` 时可向 `ctx.events` 派发 `formula:engine:ready`。
 * 实际解析/求值仍由 `@flexsheet/formula` 模块 API 提供。
 */
export class FormulaEnginePlugin extends PluginBase {
  readonly name = "flexsheet.formula";

  override install(ctx: PluginContext): void {
    ctx.events.emit("formula:engine:ready", { workbook: ctx.workbook });
  }
}
