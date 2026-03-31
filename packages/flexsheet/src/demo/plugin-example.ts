/**
 * 示例：自定义插件 — 在 Workspace 上注册、启用/禁用、与其他插件通过 events / register 协作。
 *
 * ```ts
 * import { PluginBase, type PluginContext, type IPlugin } from "flexsheet";
 *
 * class HelloPlugin extends PluginBase implements IPlugin {
 *   readonly name = "hello";
 *   install(ctx: PluginContext): void {
 *     ctx.events.on("sheet:ready", () => {
 *       console.info("Hello from plugin");
 *     });
 *   }
 * }
 *
 * const fs = new FlexSheet({ container: el });
 * fs.workspace.use(new HelloPlugin());
 * ```
 */
export {};
