/**
 * 判断 keydown 是否来自应保留浏览器默认行为的可编辑控件（编辑栏、Ribbon 输入框等）。
 * 用于避免在捕获阶段拦截 Ctrl+V / Ctrl+Z 等快捷键。
 */
export function isEditableKeydownTarget(ev: KeyboardEvent): boolean {
  const t = ev.target;
  if (t === null) {
    return false;
  }
  const n = t as Node;
  const el = n.nodeType === Node.TEXT_NODE ? (n as Text).parentElement : (n as HTMLElement);
  if (el === null) {
    return false;
  }
  if (el instanceof HTMLInputElement || el instanceof HTMLTextAreaElement) {
    return true;
  }
  if (el.isContentEditable) {
    return true;
  }
  return el.closest("input, textarea, [contenteditable='true']") !== null;
}
