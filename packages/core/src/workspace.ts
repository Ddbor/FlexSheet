import type { CommandManager } from "./command-manager.js";
import { CommandManager as CommandManagerImpl } from "./command-manager.js";
import type { EventEmitter } from "./event-emitter.js";
import { EventEmitter as EventEmitterImpl } from "./event-emitter.js";
import type { IPlugin, PluginContext } from "./plugin-types.js";
import type { UIRegistry } from "./ui-registry.js";
import { UIRegistry as UIRegistryImpl } from "./ui-registry.js";
import type { Workbook } from "./workbook.js";

class PluginContextImpl implements PluginContext {
  private readonly registry = new Map<string, unknown>();

  constructor(
    readonly workspace: Workspace,
    readonly workbook: Workbook,
    readonly commands: CommandManager,
    readonly events: EventEmitter,
    readonly ui: UIRegistry,
  ) {}

  register<T>(key: string, value: T): void {
    this.registry.set(key, value);
  }

  get<T>(key: string): T | undefined {
    return this.registry.get(key) as T | undefined;
  }

  has(key: string): boolean {
    return this.registry.has(key);
  }

  unregister(key: string): void {
    this.registry.delete(key);
  }

  clearRegistry(): void {
    this.registry.clear();
  }
}

type PluginEntry = {
  readonly plugin: IPlugin;
  active: boolean;
};

/**
 * 根容器：持有 Workbook、命令、事件、UI 注册表，并管理插件生命周期。
 */
export class Workspace {
  readonly workbook: Workbook;
  readonly commands: CommandManager;
  readonly events: EventEmitter;
  readonly ui: UIRegistry;

  private readonly ctx: PluginContextImpl;
  private readonly plugins = new Map<string, PluginEntry>();

  constructor(workbook: Workbook) {
    this.workbook = workbook;
    this.commands = new CommandManagerImpl();
    this.events = new EventEmitterImpl();
    this.ui = new UIRegistryImpl();
    this.ctx = new PluginContextImpl(this, workbook, this.commands, this.events, this.ui);
  }

  /** 创建可传给插件的上下文（同一 Workspace 共享服务表）。 */
  get pluginContext(): PluginContext {
    return this.ctx;
  }

  use(plugin: IPlugin): void {
    if (this.plugins.has(plugin.name)) {
      throw new Error(`Plugin "${plugin.name}" is already registered`);
    }
    plugin.install(this.ctx);
    this.plugins.set(plugin.name, { plugin, active: false });
    this.enable(plugin.name);
  }

  enable(name: string): boolean {
    const entry = this.plugins.get(name);
    if (entry === undefined) {
      return false;
    }
    if (entry.active) {
      return true;
    }
    entry.plugin.activate();
    entry.active = true;
    return true;
  }

  disable(name: string): boolean {
    const entry = this.plugins.get(name);
    if (entry === undefined) {
      return false;
    }
    if (!entry.active) {
      return true;
    }
    entry.plugin.deactivate();
    entry.active = false;
    return true;
  }

  uninstall(name: string): boolean {
    const entry = this.plugins.get(name);
    if (entry === undefined) {
      return false;
    }
    if (entry.active) {
      entry.plugin.deactivate();
    }
    entry.plugin.destroy();
    this.plugins.delete(name);
    return true;
  }

  isPluginActive(name: string): boolean {
    return this.plugins.get(name)?.active === true;
  }

  listPlugins(): readonly string[] {
    return [...this.plugins.keys()];
  }

  getPlugin(name: string): IPlugin | undefined {
    return this.plugins.get(name)?.plugin;
  }

  /** 销毁 Workspace：按注册逆序 destroy 插件并清空服务表。 */
  destroy(): void {
    const names = [...this.plugins.keys()].reverse();
    for (const name of names) {
      this.uninstall(name);
    }
    this.ctx.clearRegistry();
    this.events.removeAllListeners();
    this.commands.clear();
  }
}
