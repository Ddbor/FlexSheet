/**
 * 轻量类型化事件总线，供插件与宿主通信。
 */

export type EventHandler<T = unknown> = (payload: T) => void;

export class EventEmitter {
  private readonly listeners = new Map<string, Set<EventHandler>>();

  on<T = unknown>(event: string, handler: EventHandler<T>): () => void {
    let set = this.listeners.get(event);
    if (set === undefined) {
      set = new Set();
      this.listeners.set(event, set);
    }
    set.add(handler as EventHandler);
    return () => {
      this.off(event, handler as EventHandler);
    };
  }

  off(event: string, handler: EventHandler): void {
    const set = this.listeners.get(event);
    if (set === undefined) {
      return;
    }
    set.delete(handler);
    if (set.size === 0) {
      this.listeners.delete(event);
    }
  }

  emit<T = unknown>(event: string, payload?: T): void {
    const set = this.listeners.get(event);
    if (set === undefined) {
      return;
    }
    for (const fn of set) {
      fn(payload as T);
    }
  }

  /** 移除某事件下全部监听。 */
  removeAllListeners(event?: string): void {
    if (event === undefined) {
      this.listeners.clear();
      return;
    }
    this.listeners.delete(event);
  }
}
