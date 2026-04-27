const DIALOG_DRAG_STYLE_ID = "fs-dialog-drag-style";

function ensureDialogDragStyle(): void {
  if (document.getElementById(DIALOG_DRAG_STYLE_ID) !== null) {
    return;
  }
  const s = document.createElement("style");
  s.id = DIALOG_DRAG_STYLE_ID;
  s.textContent = `
.fs-dialog-drag-handle {
  cursor: grab;
  user-select: none;
  -webkit-user-select: none;
  touch-action: none;
}
.fs-dialog-drag-handle:active {
  cursor: grabbing;
}
`;
  document.head.appendChild(s);
}

/**
 * 在「标题/标题栏」上按下拖动，使 `panel` 在视口内以 `position: fixed` 平移；松手时约束在窗口内（与查找替换等对话框一致）。
 */
export function attachDraggableDialogPanel(panel: HTMLElement, dragHandle: HTMLElement): void {
  ensureDialogDragStyle();
  dragHandle.classList.add("fs-dialog-drag-handle");
  if (dragHandle.title === "") {
    dragHandle.title = "按住可拖动";
  }
  let drag: { x0: number; y0: number; l0: number; t0: number } | null = null;
  const clampPanel = (): void => {
    const w = panel.offsetWidth;
    const h = panel.offsetHeight;
    const left = Number.parseFloat(panel.style.left) || 0;
    const top = Number.parseFloat(panel.style.top) || 0;
    const maxL = Math.max(4, window.innerWidth - w - 4);
    const maxT = Math.max(4, window.innerHeight - h - 4);
    panel.style.left = `${Math.min(Math.max(4, left), maxL)}px`;
    panel.style.top = `${Math.min(Math.max(4, top), maxT)}px`;
  };
  dragHandle.addEventListener("pointerdown", (ev) => {
    if (ev.button !== 0) {
      return;
    }
    const r = panel.getBoundingClientRect();
    panel.style.position = "fixed";
    panel.style.left = `${r.left}px`;
    panel.style.top = `${r.top}px`;
    panel.style.margin = "0";
    drag = { x0: ev.clientX, y0: ev.clientY, l0: r.left, t0: r.top };
    try {
      dragHandle.setPointerCapture(ev.pointerId);
    } catch {
      /* ignore */
    }
    const onMove = (e: PointerEvent): void => {
      if (drag === null) {
        return;
      }
      const l = drag.l0 + (e.clientX - drag.x0);
      const t = drag.t0 + (e.clientY - drag.y0);
      panel.style.left = `${l}px`;
      panel.style.top = `${t}px`;
    };
    const onUp = (e: PointerEvent): void => {
      if (e.pointerId !== undefined) {
        try {
          dragHandle.releasePointerCapture(e.pointerId);
        } catch {
          /* ignore */
        }
      }
      drag = null;
      document.removeEventListener("pointermove", onMove);
      document.removeEventListener("pointerup", onUp);
      document.removeEventListener("pointercancel", onUp);
      clampPanel();
    };
    document.addEventListener("pointermove", onMove);
    document.addEventListener("pointerup", onUp);
    document.addEventListener("pointercancel", onUp);
    ev.preventDefault();
  });
}
