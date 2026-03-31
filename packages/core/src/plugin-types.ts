import type { CommandManager } from "./command-manager.js";
import type { EventEmitter } from "./event-emitter.js";
import type { UIRegistry } from "./ui-registry.js";
import type { Workbook } from "./workbook.js";

/**
 * 插件上下文：install 阶段注入，可注册/读取服务、访问命令与事件。
 */
export interface PluginContext {
  readonly workspace: import("./workspace.js").Workspace;
  readonly workbook: Workbook;
  readonly commands: CommandManager;
  readonly events: EventEmitter;
  readonly ui: UIRegistry;
  register<T>(key: string, value: T): void;
  get<T>(key: string): T | undefined;
  has(key: string): boolean;
  unregister(key: string): void;
}

/**
 * 所有功能插件须实现此接口；生命周期由 Workspace 调度。
 */
export interface IPlugin {
  readonly name: string;
  install(ctx: PluginContext): void;
  activate(): void;
  deactivate(): void;
  destroy(): void;
}

/** 可选基类：子类按需覆盖。 */
export abstract class PluginBase implements IPlugin {
  abstract readonly name: string;

  install(_ctx: PluginContext): void {}

  activate(): void {}

  deactivate(): void {}

  destroy(): void {}
}
