import { Workbook, Worksheet, type SelectionRange } from "@flexsheet/core";
import { recalcWorksheet } from "@flexsheet/formula";

import type { FlexSheet } from "../flex-sheet.js";

/** 与 `CanvasRenderer` 内 VIEW_ZOOM_MIN / MAX 对应的百分比。 */
const ZOOM_SLIDER_PCT_MIN = 25;
const ZOOM_SLIDER_PCT_MAX = 400;

/** 分隔手柄宽度（与 CSS `.fs-excel-sheet-bar__grip` width 一致）。 */
const SHEET_BAR_GRIP_PX = 8;
/** 与 `.fs-excel-sheet-bar__right` 的 `gap` 一致（grip 与滚动条区间距）。 */
const SHEET_BAR_RIGHT_FLEX_GAP_PX = 5;
const DEFAULT_GRID_HSCROLL_PANE_PX = 168;
const HSCROLL_MIN_THUMB_PX = 22;
const HSCROLL_STEP_PX = 56;

interface SelectionStats {
  readonly nonEmptyCount: number;
  readonly numericCount: number;
  readonly sum: number;
}

export interface MountExcelBottomBarOptions {
  readonly container: HTMLElement;
  readonly flexSheet: FlexSheet;
}

function computeSelectionStats(sheet: Worksheet, range: SelectionRange): SelectionStats {
  let nonEmptyCount = 0;
  let numericCount = 0;
  let sum = 0;
  for (let r = range.startRow; r <= range.endRow; r++) {
    for (let c = range.startCol; c <= range.endCol; c++) {
      const cell = sheet.getCell(r, c);
      const value = cell.value;
      if (value !== null && !(typeof value === "string" && value.trim().length === 0)) {
        nonEmptyCount++;
      }
      if (typeof value === "number" && Number.isFinite(value)) {
        numericCount++;
        sum += value;
      }
    }
  }
  return { nonEmptyCount, numericCount, sum };
}

function formatStatNumber(value: number): string {
  if (!Number.isFinite(value)) {
    return "0";
  }
  return new Intl.NumberFormat("zh-CN", {
    maximumFractionDigits: 8,
  }).format(value);
}

function nextDefaultSheetName(wb: Workbook): string {
  let max = 0;
  for (let i = 0; i < wb.sheetCount; i++) {
    const sh = wb.getSheet(i);
    if (sh === undefined) {
      continue;
    }
    const m = /^Sheet(\d+)$/i.exec(sh.name);
    if (m !== null) {
      max = Math.max(max, Number.parseInt(m[1], 10));
    }
  }
  return `Sheet${max + 1}`;
}

/**
 * 挂载类 Excel 底部「工作表标签栏 + 状态栏」；与 `Workbook` / `CanvasRenderer` 通过 `FlexSheet` 联动。
 */
export function mountExcelBottomBar(options: MountExcelBottomBarOptions): () => void {
  const { container, flexSheet } = options;
  container.classList.add("fs-excel-bottom");
  container.replaceChildren();

  const sheetBar = document.createElement("div");
  sheetBar.className = "fs-excel-sheet-bar";
  sheetBar.setAttribute("role", "toolbar");
  sheetBar.setAttribute("aria-label", "工作表标签");

  const sheetBarLeft = document.createElement("div");
  sheetBarLeft.className = "fs-excel-sheet-bar__left";

  const sheetBarRight = document.createElement("div");
  sheetBarRight.className = "fs-excel-sheet-bar__right";

  const grip = document.createElement("div");
  grip.className = "fs-excel-sheet-bar__grip";
  grip.setAttribute("role", "separator");
  grip.setAttribute("aria-orientation", "vertical");
  grip.setAttribute("aria-label", "调整横向滚动条区域宽度");
  grip.innerHTML =
    '<span class="fs-excel-sheet-bar__grip-dot"></span><span class="fs-excel-sheet-bar__grip-dot"></span><span class="fs-excel-sheet-bar__grip-dot"></span>';

  const gridHScroll = document.createElement("div");
  gridHScroll.className = "fs-excel-sheet-bar__grid-hscroll";

  const hBtnLeft = document.createElement("button");
  hBtnLeft.type = "button";
  hBtnLeft.className = "fs-excel-sheet-bar__hbtn";
  hBtnLeft.setAttribute("aria-label", "向左滚动表格");
  hBtnLeft.textContent = "◀";

  const hTrack = document.createElement("div");
  hTrack.className = "fs-excel-sheet-bar__h-track";

  const hThumb = document.createElement("div");
  hThumb.className = "fs-excel-sheet-bar__h-thumb";
  hThumb.setAttribute("role", "slider");
  hThumb.setAttribute("aria-orientation", "horizontal");
  hThumb.tabIndex = 0;

  const hBtnRight = document.createElement("button");
  hBtnRight.type = "button";
  hBtnRight.className = "fs-excel-sheet-bar__hbtn";
  hBtnRight.setAttribute("aria-label", "向右滚动表格");
  hBtnRight.textContent = "▶";

  hTrack.appendChild(hThumb);
  gridHScroll.append(hBtnLeft, hTrack, hBtnRight);
  sheetBarRight.append(grip, gridHScroll);

  const btnScrollLeft = document.createElement("button");
  btnScrollLeft.type = "button";
  btnScrollLeft.className = "fs-excel-sheet-bar__scroll fs-excel-sheet-bar__scroll--left";
  btnScrollLeft.setAttribute("aria-label", "向左滚动工作表标签");
  btnScrollLeft.textContent = "◀";

  const stripOuter = document.createElement("div");
  stripOuter.className = "fs-excel-sheet-bar__strip-outer";

  const stripScroll = document.createElement("div");
  stripScroll.className = "fs-excel-sheet-bar__strip-scroll";

  const tabsRow = document.createElement("div");
  tabsRow.className = "fs-excel-sheet-bar__tabs";

  stripScroll.appendChild(tabsRow);
  stripOuter.appendChild(stripScroll);

  const btnScrollRight = document.createElement("button");
  btnScrollRight.type = "button";
  btnScrollRight.className = "fs-excel-sheet-bar__scroll";
  btnScrollRight.setAttribute("aria-label", "向右滚动工作表标签");
  btnScrollRight.textContent = "▶";

  const btnAdd = document.createElement("button");
  btnAdd.type = "button";
  btnAdd.className = "fs-excel-sheet-bar__add";
  btnAdd.setAttribute("aria-label", "新建工作表");
  btnAdd.textContent = "+";

  sheetBarLeft.append(btnScrollLeft, btnScrollRight, stripOuter);
  sheetBar.append(sheetBarLeft, sheetBarRight);

  const statusBar = document.createElement("div");
  statusBar.className = "fs-excel-status-bar";
  statusBar.setAttribute("role", "status");

  const statusLeft = document.createElement("span");
  statusLeft.className = "fs-excel-status-bar__msg";
  statusLeft.textContent = "就绪";

  const statusStats = document.createElement("span");
  statusStats.className = "fs-excel-status-bar__stats";

  const zoomWrap = document.createElement("div");
  zoomWrap.className = "fs-excel-status-bar__zoom-wrap";

  const zoomOutBtn = document.createElement("button");
  zoomOutBtn.type = "button";
  zoomOutBtn.className = "fs-excel-status-bar__zoom-icon";
  zoomOutBtn.setAttribute("aria-label", "缩小");
  zoomOutBtn.textContent = "−";

  const zoomRange = document.createElement("input");
  zoomRange.type = "range";
  zoomRange.className = "fs-excel-status-bar__zoom-range";
  zoomRange.min = String(ZOOM_SLIDER_PCT_MIN);
  zoomRange.max = String(ZOOM_SLIDER_PCT_MAX);
  zoomRange.step = "1";
  zoomRange.setAttribute("aria-label", "缩放比例");

  const zoomInBtn = document.createElement("button");
  zoomInBtn.type = "button";
  zoomInBtn.className = "fs-excel-status-bar__zoom-icon";
  zoomInBtn.setAttribute("aria-label", "放大");
  zoomInBtn.textContent = "+";

  const zoomReadout = document.createElement("span");
  zoomReadout.className = "fs-excel-status-bar__zoom-readout";

  zoomWrap.append(zoomOutBtn, zoomRange, zoomInBtn, zoomReadout);
  statusBar.append(statusLeft, statusStats, zoomWrap);
  container.append(sheetBar, statusBar);

  const renderer = flexSheet.getRenderer();

  const clampPct = (pct: number): number =>
    Math.max(ZOOM_SLIDER_PCT_MIN, Math.min(ZOOM_SLIDER_PCT_MAX, Math.round(pct)));

  const syncZoomUi = (): void => {
    const pct = clampPct(renderer.getViewZoom() * 100);
    zoomReadout.textContent = `${pct}%`;
    zoomRange.value = String(pct);
  };

  const syncSelectionStats = (): void => {
    const sheet = flexSheet.workbook.getActiveSheet();
    if (sheet === undefined) {
      statusStats.textContent = "";
      return;
    }
    const range = flexSheet.selection.getNormalizedRange();
    const stats = computeSelectionStats(sheet, range);
    if (stats.nonEmptyCount <= 1) {
      statusStats.textContent = "";
      return;
    }
    const avg = stats.numericCount > 0 ? stats.sum / stats.numericCount : 0;
    statusStats.textContent = `平均值: ${formatStatNumber(avg)}    计数: ${formatStatNumber(stats.nonEmptyCount)}    求和: ${formatStatNumber(stats.sum)}`;
  };

  syncZoomUi();
  syncSelectionStats();

  let gridHScrollPanePx = DEFAULT_GRID_HSCROLL_PANE_PX;

  /** 标签栏内容区宽度（已扣除左右 padding，与 flex 子项可用宽度一致）。 */
  const sheetBarContentWidthPx = (): number => {
    const cs = getComputedStyle(sheetBar);
    const pl = Number.parseFloat(cs.paddingLeft) || 0;
    const pr = Number.parseFloat(cs.paddingRight) || 0;
    return Math.max(0, sheetBar.clientWidth - pl - pr);
  };

  /** 右侧总宽 = grip + gap + 滚动条区，不得超过内容区宽度（避免拖出栏外）。 */
  const maxGridScrollPanePx = (): number =>
    Math.max(0, sheetBarContentWidthPx() - SHEET_BAR_GRIP_PX - SHEET_BAR_RIGHT_FLEX_GAP_PX);

  const applyGridScrollPaneStyle = (): void => {
    gridHScrollPanePx = Math.max(0, Math.min(gridHScrollPanePx, maxGridScrollPanePx()));
    if (gridHScrollPanePx <= 0) {
      gridHScroll.classList.add("fs-excel-sheet-bar__grid-hscroll--collapsed");
      gridHScroll.style.width = "0px";
      gridHScroll.style.minWidth = "0px";
    } else {
      gridHScroll.classList.remove("fs-excel-sheet-bar__grid-hscroll--collapsed");
      gridHScroll.style.width = `${gridHScrollPanePx}px`;
      gridHScroll.style.minWidth = `${gridHScrollPanePx}px`;
    }
  };

  const applyCanvasScrollX = (nextX: number): void => {
    const y = renderer.getScroll().scrollY;
    renderer.setScroll(nextX, y);
    renderer.ensureScrollClamped();
    flexSheet.refresh();
  };

  const syncGridHScrollbar = (): void => {
    applyGridScrollPaneStyle();
    if (gridHScrollPanePx <= 0) {
      return;
    }

    const m = renderer.getHorizontalScrollMetrics();
    if (m === null) {
      return;
    }

    const { scrollX, maxScrollX, scrollViewportW, contentScrollWidth } = m;
    const trackW = hTrack.clientWidth;
    if (trackW <= 0) {
      return;
    }

    if (maxScrollX <= 0 || contentScrollWidth <= scrollViewportW + 0.5) {
      hThumb.style.width = `${trackW}px`;
      hThumb.style.left = "0px";
      hBtnLeft.disabled = true;
      hBtnRight.disabled = true;
      hThumb.setAttribute("aria-valuemin", "0");
      hThumb.setAttribute("aria-valuemax", "0");
      hThumb.setAttribute("aria-valuenow", "0");
      return;
    }

    const ratio = scrollViewportW / Math.max(contentScrollWidth, 1);
    let thumbW = Math.max(HSCROLL_MIN_THUMB_PX, ratio * trackW);
    thumbW = Math.min(thumbW, trackW);
    const span = Math.max(1e-6, trackW - thumbW);
    const t = maxScrollX > 0 ? scrollX / maxScrollX : 0;
    const left = span * t;
    hThumb.style.width = `${thumbW}px`;
    hThumb.style.left = `${left}px`;

    hBtnLeft.disabled = scrollX <= 0.5;
    hBtnRight.disabled = scrollX >= maxScrollX - 0.5;

    hThumb.setAttribute("aria-valuemin", "0");
    hThumb.setAttribute("aria-valuemax", String(Math.round(maxScrollX)));
    hThumb.setAttribute("aria-valuenow", String(Math.round(scrollX)));
  };

  hBtnLeft.addEventListener("click", () => {
    applyCanvasScrollX(renderer.getScroll().scrollX - HSCROLL_STEP_PX);
  });
  hBtnRight.addEventListener("click", () => {
    applyCanvasScrollX(renderer.getScroll().scrollX + HSCROLL_STEP_PX);
  });

  hTrack.addEventListener("pointerdown", (ev) => {
    if (ev.target !== hTrack) {
      return;
    }
    const m = renderer.getHorizontalScrollMetrics();
    if (m === null || m.maxScrollX <= 0) {
      return;
    }
    const rect = hTrack.getBoundingClientRect();
    const x = ev.clientX - rect.left;
    const tw = hThumb.offsetWidth;
    const span = rect.width - tw;
    if (span <= 0) {
      return;
    }
    const cx = Math.max(tw / 2, Math.min(x, rect.width - tw / 2));
    const ratio = (cx - tw / 2) / span;
    applyCanvasScrollX(ratio * m.maxScrollX);
  });

  let thumbDrag: {
    readonly pointerId: number;
    readonly startClientX: number;
    readonly scroll0: number;
    readonly span: number;
    readonly maxS: number;
  } | null = null;

  hThumb.addEventListener("pointerdown", (ev) => {
    ev.stopPropagation();
    const m = renderer.getHorizontalScrollMetrics();
    if (m === null || m.maxScrollX <= 0) {
      return;
    }
    const tw = hThumb.offsetWidth;
    const span = hTrack.clientWidth - tw;
    if (span <= 0) {
      return;
    }
    thumbDrag = {
      pointerId: ev.pointerId,
      startClientX: ev.clientX,
      scroll0: m.scrollX,
      span,
      maxS: m.maxScrollX,
    };
    hThumb.setPointerCapture(ev.pointerId);
    hThumb.classList.add("fs-excel-sheet-bar__h-thumb--dragging");
  });

  hThumb.addEventListener("pointermove", (ev) => {
    if (thumbDrag === null || ev.pointerId !== thumbDrag.pointerId) {
      return;
    }
    const dx = ev.clientX - thumbDrag.startClientX;
    applyCanvasScrollX(thumbDrag.scroll0 + (dx / thumbDrag.span) * thumbDrag.maxS);
  });

  const endThumbDrag = (ev: PointerEvent): void => {
    if (thumbDrag === null || ev.pointerId !== thumbDrag.pointerId) {
      return;
    }
    try {
      hThumb.releasePointerCapture(ev.pointerId);
    } catch {
      /* ignore */
    }
    hThumb.classList.remove("fs-excel-sheet-bar__h-thumb--dragging");
    thumbDrag = null;
  };
  hThumb.addEventListener("pointerup", endThumbDrag);
  hThumb.addEventListener("pointercancel", endThumbDrag);

  let gripDrag: { readonly pointerId: number; readonly startX: number; readonly startPane: number } | null =
    null;

  grip.addEventListener("pointerdown", (ev) => {
    ev.preventDefault();
    gripDrag = {
      pointerId: ev.pointerId,
      startX: ev.clientX,
      startPane: gridHScrollPanePx,
    };
    grip.setPointerCapture(ev.pointerId);
    grip.classList.add("fs-excel-sheet-bar__grip--dragging");
  });

  const onGripPointerMove = (ev: PointerEvent): void => {
    if (gripDrag === null || ev.pointerId !== gripDrag.pointerId) {
      return;
    }
    const dx = ev.clientX - gripDrag.startX;
    gridHScrollPanePx = gripDrag.startPane - dx;
    syncGridHScrollbar();
  };

  const endGripDrag = (ev: PointerEvent): void => {
    if (gripDrag === null || ev.pointerId !== gripDrag.pointerId) {
      return;
    }
    try {
      grip.releasePointerCapture(ev.pointerId);
    } catch {
      /* ignore */
    }
    grip.classList.remove("fs-excel-sheet-bar__grip--dragging");
    gripDrag = null;
  };
  grip.addEventListener("pointermove", onGripPointerMove);
  grip.addEventListener("pointerup", endGripDrag);
  grip.addEventListener("pointercancel", endGripDrag);

  const updateScrollArrows = (): void => {
    const el = stripScroll;
    const max = el.scrollWidth - el.clientWidth;
    btnScrollLeft.disabled = el.scrollLeft <= 1;
    btnScrollRight.disabled = max <= 1 || el.scrollLeft >= max - 1;
  };

  const scrollTabs = (delta: number): void => {
    stripScroll.scrollBy({ left: delta, behavior: "smooth" });
    window.setTimeout(updateScrollArrows, 120);
  };

  btnScrollLeft.addEventListener("click", () => {
    scrollTabs(-72);
  });
  btnScrollRight.addEventListener("click", () => {
    scrollTabs(72);
  });
  stripScroll.addEventListener("scroll", updateScrollArrows);

  let menuEl: HTMLDivElement | null = null;
  let menuDown: ((e: MouseEvent) => void) | null = null;

  const closeMenu = (): void => {
    if (menuEl !== null) {
      menuEl.remove();
      menuEl = null;
    }
    if (menuDown !== null) {
      document.removeEventListener("mousedown", menuDown, true);
      menuDown = null;
    }
  };

  const openSheetContextMenu = (clientX: number, clientY: number, sheetIndex: number): void => {
    closeMenu();
    const wb = flexSheet.workbook;
    const sheet = wb.getSheet(sheetIndex);
    if (sheet === undefined) {
      return;
    }

    const menu = document.createElement("div");
    menu.className = "fs-excel-sheet-menu";
    menu.setAttribute("role", "menu");

    const itemRename = document.createElement("button");
    itemRename.type = "button";
    itemRename.className = "fs-excel-sheet-menu__item";
    itemRename.textContent = "重命名";
    itemRename.addEventListener("click", () => {
      closeMenu();
      const raw = window.prompt("工作表名称", sheet.name);
      if (raw !== null) {
        const t = raw.trim();
        if (t.length > 0) {
          sheet.setName(t);
        }
      }
    });

    const itemDelete = document.createElement("button");
    itemDelete.type = "button";
    itemDelete.className = "fs-excel-sheet-menu__item";
    itemDelete.textContent = "删除";
    itemDelete.disabled = wb.sheetCount <= 1;
    itemDelete.addEventListener("click", () => {
      closeMenu();
      if (wb.sheetCount <= 1) {
        return;
      }
      if (!window.confirm(`确定删除工作表「${sheet.name}」？`)) {
        return;
      }
      wb.removeSheetAt(sheetIndex);
    });

    menu.append(itemRename, itemDelete);
    document.body.appendChild(menu);
    menuEl = menu;

    const pad = 4;
    let x = clientX;
    let y = clientY;
    const mw = menu.offsetWidth || 120;
    const mh = menu.offsetHeight || 64;
    if (x + mw > window.innerWidth - pad) {
      x = window.innerWidth - mw - pad;
    }
    if (y + mh > window.innerHeight - pad) {
      y = window.innerHeight - mh - pad;
    }
    menu.style.left = `${Math.max(pad, x)}px`;
    menu.style.top = `${Math.max(pad, y)}px`;

    menuDown = (e: MouseEvent): void => {
      if (!menu.contains(e.target as Node)) {
        closeMenu();
      }
    };
    window.setTimeout(() => {
      document.addEventListener("mousedown", menuDown!, true);
    }, 0);
  };

  const renderTabs = (): void => {
    tabsRow.replaceChildren();
    const wb = flexSheet.workbook;
    for (let i = 0; i < wb.sheetCount; i++) {
      const sh = wb.getSheet(i);
      if (sh === undefined) {
        continue;
      }
      const tab = document.createElement("button");
      tab.type = "button";
      tab.className = "fs-excel-sheet-tab";
      if (i === wb.activeSheetIndex) {
        tab.classList.add("fs-excel-sheet-tab--active");
        tab.setAttribute("aria-current", "true");
      } else {
        tab.removeAttribute("aria-current");
      }
      tab.textContent = sh.name;
      tab.addEventListener("click", () => {
        wb.activeSheetIndex = i;
        requestAnimationFrame(() => {
          tab.scrollIntoView({ block: "nearest", inline: "nearest" });
        });
      });
      tab.addEventListener("contextmenu", (ev) => {
        ev.preventDefault();
        openSheetContextMenu(ev.clientX, ev.clientY, i);
      });
      tabsRow.appendChild(tab);
    }
    tabsRow.appendChild(btnAdd);
    requestAnimationFrame(() => {
      updateScrollArrows();
    });
  };

  btnAdd.addEventListener("click", () => {
    const wb = flexSheet.workbook;
    const name = nextDefaultSheetName(wb);
    const sh = new Worksheet(name, 200, 32);
    wb.addSheet(sh);
    wb.activeSheetIndex = wb.sheetCount - 1;
    recalcWorksheet(sh);
  });

  zoomRange.addEventListener("input", () => {
    const v = Number.parseInt(zoomRange.value, 10);
    if (!Number.isFinite(v)) {
      return;
    }
    renderer.setViewZoom(clampPct(v) / 100);
    flexSheet.refresh();
  });

  zoomOutBtn.addEventListener("click", () => {
    renderer.zoomOut();
    flexSheet.refresh();
  });
  zoomInBtn.addEventListener("click", () => {
    renderer.zoomIn();
    flexSheet.refresh();
  });

  let unsubWb: () => void = () => {};

  const bindWorkbookSubscription = (): void => {
    unsubWb();
    unsubWb = flexSheet.workbook.subscribe(() => {
      renderTabs();
      syncGridHScrollbar();
    });
  };

  const unsubWorkbookReplaced = flexSheet.subscribeWorkbookReplaced(() => {
    bindWorkbookSubscription();
    renderTabs();
    syncGridHScrollbar();
  });
  const unsubZoom = renderer.subscribeViewZoom(() => {
    syncZoomUi();
    syncGridHScrollbar();
  });
  const unsubScroll = renderer.subscribeScroll(() => {
    syncGridHScrollbar();
  });
  const unsubStats = flexSheet.subscribeFormattingChrome(() => {
    syncSelectionStats();
  });

  const ro = typeof ResizeObserver !== "undefined" ? new ResizeObserver(updateScrollArrows) : null;
  ro?.observe(stripScroll);

  const roSheetBar =
    typeof ResizeObserver !== "undefined"
      ? new ResizeObserver(() => {
          syncGridHScrollbar();
        })
      : null;
  roSheetBar?.observe(sheetBar);

  return (): void => {
    closeMenu();
    unsubWb();
    unsubWorkbookReplaced();
    unsubZoom();
    unsubScroll();
    unsubStats();
    ro?.disconnect();
    roSheetBar?.disconnect();
    container.replaceChildren();
    container.classList.remove("fs-excel-bottom");
  };
}
