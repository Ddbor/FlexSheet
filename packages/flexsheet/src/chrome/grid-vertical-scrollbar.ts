import type { FlexSheet } from "../flex-sheet.js";

const VSCROLL_STEP_PX = 48;
const VSCROLL_MIN_THUMB_PX = 22;

export interface MountGridVerticalScrollbarOptions {
  /** 竖条容器（将铺满画布区高度，与 canvas 宿主并列）。 */
  readonly container: HTMLElement;
  readonly flexSheet: FlexSheet;
}

/**
 * 在画布右侧挂载与 `CanvasRenderer.scrollY` 联动的假竖向滚动条（白底灰边按钮 + 轨道 + 滑块）。
 */
export function mountGridVerticalScrollbar(options: MountGridVerticalScrollbarOptions): () => void {
  const { container, flexSheet } = options;
  const renderer = flexSheet.getRenderer();

  container.classList.add("fs-grid-vscroll");
  container.replaceChildren();
  container.setAttribute("role", "scrollbar");
  container.setAttribute("aria-orientation", "vertical");

  const vBtnUp = document.createElement("button");
  vBtnUp.type = "button";
  vBtnUp.className = "fs-grid-vscroll__btn";
  vBtnUp.setAttribute("aria-label", "向上滚动");
  vBtnUp.textContent = "▲";

  const vTrack = document.createElement("div");
  vTrack.className = "fs-grid-vscroll__track";

  const vThumb = document.createElement("div");
  vThumb.className = "fs-grid-vscroll__thumb";
  vThumb.setAttribute("role", "slider");
  vThumb.setAttribute("aria-orientation", "vertical");
  vThumb.tabIndex = 0;

  const vBtnDown = document.createElement("button");
  vBtnDown.type = "button";
  vBtnDown.className = "fs-grid-vscroll__btn";
  vBtnDown.setAttribute("aria-label", "向下滚动");
  vBtnDown.textContent = "▼";

  vTrack.appendChild(vThumb);
  container.append(vBtnUp, vTrack, vBtnDown);

  const applyCanvasScrollY = (nextY: number): void => {
    const x = renderer.getScroll().scrollX;
    renderer.setScroll(x, nextY);
    renderer.ensureScrollClamped();
    flexSheet.refresh();
  };

  const syncVScrollbar = (): void => {
    const m = renderer.getVerticalScrollMetrics();
    if (m === null) {
      return;
    }

    const trackH = vTrack.clientHeight;
    if (trackH <= 0) {
      return;
    }

    const { scrollY, maxScrollY, scrollViewportH, contentScrollHeight } = m;

    if (maxScrollY <= 0 || contentScrollHeight <= scrollViewportH + 0.5) {
      vThumb.style.height = `${trackH}px`;
      vThumb.style.top = "0px";
      vBtnUp.disabled = true;
      vBtnDown.disabled = true;
      vThumb.setAttribute("aria-valuemin", "0");
      vThumb.setAttribute("aria-valuemax", "0");
      vThumb.setAttribute("aria-valuenow", "0");
      return;
    }

    const ratio = scrollViewportH / Math.max(contentScrollHeight, 1);
    let thumbH = Math.max(VSCROLL_MIN_THUMB_PX, ratio * trackH);
    thumbH = Math.min(thumbH, trackH);
    const span = Math.max(1e-6, trackH - thumbH);
    const t = maxScrollY > 0 ? scrollY / maxScrollY : 0;
    const top = span * t;
    vThumb.style.height = `${thumbH}px`;
    vThumb.style.top = `${top}px`;

    vBtnUp.disabled = scrollY <= 0.5;
    vBtnDown.disabled = scrollY >= maxScrollY - 0.5;

    vThumb.setAttribute("aria-valuemin", "0");
    vThumb.setAttribute("aria-valuemax", String(Math.round(maxScrollY)));
    vThumb.setAttribute("aria-valuenow", String(Math.round(scrollY)));
  };

  vBtnUp.addEventListener("click", () => {
    applyCanvasScrollY(renderer.getScroll().scrollY - VSCROLL_STEP_PX);
  });
  vBtnDown.addEventListener("click", () => {
    applyCanvasScrollY(renderer.getScroll().scrollY + VSCROLL_STEP_PX);
  });

  vTrack.addEventListener("pointerdown", (ev) => {
    if (ev.target !== vTrack) {
      return;
    }
    const m = renderer.getVerticalScrollMetrics();
    if (m === null || m.maxScrollY <= 0) {
      return;
    }
    const rect = vTrack.getBoundingClientRect();
    const y = ev.clientY - rect.top;
    const th = vThumb.offsetHeight;
    const span = rect.height - th;
    if (span <= 0) {
      return;
    }
    const cy = Math.max(th / 2, Math.min(y, rect.height - th / 2));
    const ratio = (cy - th / 2) / span;
    applyCanvasScrollY(ratio * m.maxScrollY);
  });

  let thumbDrag: {
    readonly pointerId: number;
    readonly startClientY: number;
    readonly scroll0: number;
    readonly span: number;
    readonly maxS: number;
  } | null = null;

  vThumb.addEventListener("pointerdown", (ev) => {
    ev.stopPropagation();
    const m = renderer.getVerticalScrollMetrics();
    if (m === null || m.maxScrollY <= 0) {
      return;
    }
    const th = vThumb.offsetHeight;
    const span = vTrack.clientHeight - th;
    if (span <= 0) {
      return;
    }
    thumbDrag = {
      pointerId: ev.pointerId,
      startClientY: ev.clientY,
      scroll0: m.scrollY,
      span,
      maxS: m.maxScrollY,
    };
    vThumb.setPointerCapture(ev.pointerId);
    vThumb.classList.add("fs-grid-vscroll__thumb--dragging");
  });

  vThumb.addEventListener("pointermove", (ev) => {
    if (thumbDrag === null || ev.pointerId !== thumbDrag.pointerId) {
      return;
    }
    const dy = ev.clientY - thumbDrag.startClientY;
    applyCanvasScrollY(thumbDrag.scroll0 + (dy / thumbDrag.span) * thumbDrag.maxS);
  });

  const endThumbDrag = (ev: PointerEvent): void => {
    if (thumbDrag === null || ev.pointerId !== thumbDrag.pointerId) {
      return;
    }
    try {
      vThumb.releasePointerCapture(ev.pointerId);
    } catch {
      /* ignore */
    }
    vThumb.classList.remove("fs-grid-vscroll__thumb--dragging");
    thumbDrag = null;
  };
  vThumb.addEventListener("pointerup", endThumbDrag);
  vThumb.addEventListener("pointercancel", endThumbDrag);

  const unsubScroll = renderer.subscribeScroll(() => {
    syncVScrollbar();
  });
  const unsubZoom = renderer.subscribeViewZoom(() => {
    syncVScrollbar();
  });
  const unsubWb = flexSheet.workbook.subscribe(() => {
    syncVScrollbar();
  });

  const roHost =
    typeof ResizeObserver !== "undefined"
      ? new ResizeObserver(() => {
          syncVScrollbar();
        })
      : null;
  const canvas = flexSheet.getCanvas();
  roHost?.observe(canvas.parentElement ?? canvas);
  roHost?.observe(container);

  requestAnimationFrame(() => {
    syncVScrollbar();
  });

  return (): void => {
    unsubScroll();
    unsubZoom();
    unsubWb();
    roHost?.disconnect();
    container.replaceChildren();
    container.classList.remove("fs-grid-vscroll");
    container.removeAttribute("role");
    container.removeAttribute("aria-orientation");
  };
}
