import {
  getPivotValueFieldCaption,
  normalizeSelectionRange,
  type ICommand,
  type PivotAggregateKind,
  type PivotValueComputed,
  type Workbook,
  type Worksheet,
} from "@flexsheet/core";
import {
  findPivotTableDefinitionAtCell,
  UpdatePivotTableLayoutCommand,
} from "./pivot-table-command.js";

export interface PivotFieldsPaneHost {
  readonly workbook: Workbook | undefined;
  readonly workspace: { readonly commands: { execute(cmd: ICommand): void } };
  refresh(): void;
  /** Canvas 挂载容器（如 `#grid-canvas-host`）；字段面板将作为其后的兄弟节点与表格局部并排。 */
  getSheetContainerElement(): HTMLElement;
}

interface SourceField {
  readonly col: number;
  readonly label: string;
}

interface ValueEntry {
  readonly id: string;
  col: number;
  aggregate: PivotAggregateKind;
  computed?: PivotValueComputed;
}

let paneStylesInjected = false;
let activePaneRoot: HTMLElement | null = null;
let paneKeydownHandler: ((ev: KeyboardEvent) => void) | null = null;
/** 当前已挂载的字段窗格所对应的透视表定义 id；用于选区同步时避免同表重复重建。 */
let openPivotFieldsPaneDefId: string | null = null;

type PivotZoneId = "filter" | "columns" | "rows" | "values";

interface PivotDragState {
  readonly kind: "fieldList" | "axisChip" | "valueChip" | "valueAxisChip";
  readonly col?: number;
  readonly entryId?: string;
  readonly fromZone?: PivotZoneId;
  readonly fromIndex: number;
}

function arrayMove<T>(arr: T[], from: number, to: number): void {
  if (from === to || from < 0 || from >= arr.length || to < 0 || to > arr.length) {
    return;
  }
  const [item] = arr.splice(from, 1);
  arr.splice(to, 0, item);
}

function dedupeColsWithinZone(cols: number[]): void {
  const seen = new Set<number>();
  for (let i = cols.length - 1; i >= 0; i--) {
    const c = cols[i]!;
    if (seen.has(c)) {
      cols.splice(i, 1);
    } else {
      seen.add(c);
    }
  }
}

function dedupeValueEntriesByCol(entries: ValueEntry[]): void {
  const seen = new Set<number>();
  for (let i = entries.length - 1; i >= 0; i--) {
    const c = entries[i]!.col;
    if (seen.has(c)) {
      entries.splice(i, 1);
    } else {
      seen.add(c);
    }
  }
}

/** 同一数据源列只能属于一个区域（与 Excel 一致）。优先级：筛选 > 列 > 行 > 值。 */
function enforceExclusivePivotColumns(
  filterCols: number[],
  columnCols: number[],
  rowCols: number[],
  valueEntries: ValueEntry[],
): void {
  dedupeColsWithinZone(filterCols);
  dedupeColsWithinZone(columnCols);
  dedupeColsWithinZone(rowCols);
  dedupeValueEntriesByCol(valueEntries);
  const used = new Set<number>();
  for (const c of filterCols) {
    used.add(c);
  }
  for (let i = columnCols.length - 1; i >= 0; i--) {
    const c = columnCols[i]!;
    if (used.has(c)) {
      columnCols.splice(i, 1);
    } else {
      used.add(c);
    }
  }
  for (let i = rowCols.length - 1; i >= 0; i--) {
    const c = rowCols[i]!;
    if (used.has(c)) {
      rowCols.splice(i, 1);
    } else {
      used.add(c);
    }
  }
  for (let i = valueEntries.length - 1; i >= 0; i--) {
    const c = valueEntries[i]!.col;
    if (used.has(c)) {
      valueEntries.splice(i, 1);
    } else {
      used.add(c);
    }
  }
}

function insertIndexFromPointer(container: HTMLElement, clientY: number): number {
  const children = [...container.children] as HTMLElement[];
  for (let i = 0; i < children.length; i++) {
    const r = children[i].getBoundingClientRect();
    const mid = r.top + r.height / 2;
    if (clientY < mid) {
      return i;
    }
  }
  return children.length;
}

function ensurePaneStyles(): void {
  if (paneStylesInjected) {
    return;
  }
  paneStylesInjected = true;
  const style = document.createElement("style");
  style.setAttribute("data-fs-pivot-fields-pane", "1");
  style.textContent = `
.fs-pivot-fields {
  box-sizing: border-box;
  background: #f3f2f1;
  border: 1px solid #d2d0ce;
  display: flex;
  flex-direction: column;
  font-family: "Segoe UI", system-ui, -apple-system, sans-serif;
  font-size: 11px;
  line-height: 1.35;
  color: #323130;
  -webkit-font-smoothing: antialiased;
}
.fs-pivot-fields--docked {
  flex: 0 0 288px;
  width: 288px;
  min-width: 0;
  max-width: min(288px, 100%);
  min-height: 0;
  align-self: stretch;
  position: relative;
  z-index: 1;
  border-radius: 0;
  box-shadow: none;
}
.fs-pivot-fields--floating {
  position: fixed;
  z-index: 12000;
  top: 72px;
  right: 12px;
  bottom: 12px;
  width: 288px;
  max-width: calc(100vw - 24px);
  border-radius: 4px;
  box-shadow: 0 4px 16px rgba(0,0,0,.12);
}
.fs-pivot-fields__head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 8px 10px;
  border-bottom: 1px solid #edebe9;
  background: #faf9f8;
  flex-shrink: 0;
}
.fs-pivot-fields__title {
  font-weight: 600;
  font-size: 12px;
  letter-spacing: -0.01em;
  color: #201f1e;
}
.fs-pivot-fields__close {
  border: none;
  background: transparent;
  cursor: pointer;
  font-size: 18px;
  line-height: 1;
  color: #605e5c;
  padding: 2px 6px;
  border-radius: 2px;
}
.fs-pivot-fields__close:hover {
  background: #edebe9;
  color: #201f1e;
}
.fs-pivot-fields__search {
  padding: 6px 10px;
  flex-shrink: 0;
}
.fs-pivot-fields__search-inner {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 3px 6px;
  background: #fff;
  border: 1px solid #8a8886;
  border-radius: 2px;
}
.fs-pivot-fields__search-inner:focus-within {
  border-color: #0078d4;
  outline: 1px solid #0078d4;
  outline-offset: -1px;
}
.fs-pivot-fields__search-icon {
  flex: 0 0 auto;
  color: #605e5c;
  display: flex;
  align-items: center;
}
.fs-pivot-fields__search-inner input {
  flex: 1;
  min-width: 0;
  border: none;
  outline: none;
  font-size: 11px;
  padding: 2px 0;
  background: transparent;
}
.fs-pivot-fields__list-wrap {
  flex: 0 1 auto;
  align-self: stretch;
  max-height: min(200px, 26vh);
  min-height: 0;
  overflow: auto;
  margin: 0 8px 6px;
  background: #fff;
  border: 1px solid #edebe9;
  border-radius: 2px;
}
.fs-pivot-fields__row {
  display: flex;
  align-items: center;
  gap: 6px;
  padding: 4px 8px;
  border-bottom: 1px solid #f3f2f1;
  cursor: grab;
  font-size: 11px;
}
.fs-pivot-fields__row:hover {
  background: #faf9f8;
}
.fs-pivot-fields__row--hidden {
  display: none;
}
.fs-pivot-fields__zones {
  flex: 1;
  min-height: 0;
  overflow: auto;
  padding: 0 8px 0;
  display: grid;
  grid-template-columns: 1fr 1fr;
  grid-template-rows: 1fr 1fr;
  gap: 6px;
  align-content: stretch;
}
.fs-pivot-fields__zone {
  background: #fff;
  border: 1px solid #e1dfdd;
  border-radius: 2px;
  padding: 0;
  min-height: 92px;
  display: flex;
  flex-direction: column;
  overflow: hidden;
  transition: border-color 0.12s ease, box-shadow 0.12s ease;
}
.fs-pivot-fields__zone--drop-target {
  border-color: #107c10;
  box-shadow: 0 0 0 1px #107c10 inset;
}
.fs-pivot-fields__zone-head {
  display: flex;
  align-items: center;
  gap: 5px;
  padding: 5px 8px 4px;
  border-bottom: 1px solid #f3f2f1;
  flex-shrink: 0;
}
.fs-pivot-fields__zone-icon {
  flex: 0 0 auto;
  width: 14px;
  height: 14px;
  color: #605e5c;
  display: flex;
  align-items: center;
  justify-content: center;
}
.fs-pivot-fields__zone-icon svg {
  display: block;
  width: 12px;
  height: 12px;
  fill: currentColor;
}
.fs-pivot-fields__zone-title {
  font-size: 10px;
  color: #605e5c;
  font-weight: 600;
  letter-spacing: 0.02em;
  flex: 1;
  min-width: 0;
}
.fs-pivot-fields__chips {
  display: flex;
  flex-direction: column;
  gap: 3px;
  padding: 4px 6px 6px;
  min-height: 40px;
  flex: 1;
  overflow: auto;
}
.fs-pivot-fields__chip {
  display: flex;
  align-items: stretch;
  gap: 0;
  background: #e8f4fc;
  border: 1px solid #b8d9f4;
  border-radius: 2px;
  font-size: 10px;
  max-width: 100%;
  min-height: 24px;
  cursor: grab;
}
.fs-pivot-fields__chip--values {
  background: #fff8e5;
  border-color: #e8c96c;
}
.fs-pivot-fields__chip--value-axis {
  background: #ede7f6;
  border-color: #b4a7d6;
}
.fs-pivot-fields__chip--value-axis .fs-pivot-fields__chip-label {
  font-weight: 600;
  color: #323130;
}
.fs-pivot-fields__chip-handle {
  flex: 0 0 14px;
  display: flex;
  align-items: center;
  justify-content: center;
  color: #b3b0ad;
  user-select: none;
  border-right: 1px solid rgba(0,0,0,.06);
}
.fs-pivot-fields__chip-handle span {
  display: block;
  width: 2px;
  height: 10px;
  background: repeating-linear-gradient(
    to bottom,
    currentColor 0 2px,
    transparent 2px 4px
  );
  border-radius: 1px;
}
.fs-pivot-fields__chip-body {
  flex: 1;
  min-width: 0;
  display: flex;
  align-items: center;
  padding: 3px 4px 3px 2px;
  gap: 2px;
}
.fs-pivot-fields__chip-label {
  flex: 1;
  min-width: 0;
  color: #201f1e;
  overflow: hidden;
  white-space: nowrap;
  text-overflow: ellipsis;
  line-height: 1.35;
}
.fs-pivot-fields__chip-x {
  flex: 0 0 auto;
  border: none;
  background: transparent;
  cursor: pointer;
  padding: 0 2px;
  font-size: 12px;
  line-height: 1;
  color: #8a8886;
  align-self: center;
}
.fs-pivot-fields__chip-x:hover {
  color: #a4262c;
}
.fs-pivot-fields__chip-info {
  flex: 0 0 auto;
  border: none;
  background: transparent;
  cursor: pointer;
  width: 14px;
  height: 14px;
  border-radius: 50%;
  font-size: 9px;
  font-weight: 700;
  font-style: italic;
  color: #0078d4;
  padding: 0;
  line-height: 1;
  display: inline-flex;
  align-items: center;
  justify-content: center;
  align-self: center;
}
.fs-pivot-fields__chip-info:hover {
  background: rgba(0, 120, 212, 0.08);
}
.fs-pivot-fields__zones-hint {
  padding: 6px 8px 8px;
  font-size: 10px;
  color: #8a8886;
  text-align: center;
  flex-shrink: 0;
  line-height: 1.35;
}
.fs-pivot-fields__warn {
  margin: 0 8px 6px;
  font-size: 10px;
  color: #a4262c;
  min-height: 14px;
  line-height: 1.35;
}
`;
  document.head.appendChild(style);
}

function newId(): string {
  return `pv-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 6)}`;
}

function collectSourceFields(
  sheet: Worksheet,
  range: ReturnType<typeof normalizeSelectionRange>,
  hasHeaders: boolean,
): SourceField[] {
  const out: SourceField[] = [];
  for (let c = range.startCol; c <= range.endCol; c++) {
    const fallback = `列${c - range.startCol + 1}`;
    const label =
      hasHeaders && typeof sheet.getCell(range.startRow, c).value === "string"
        ? String(sheet.getCell(range.startRow, c).value).trim() || fallback
        : hasHeaders
          ? String(sheet.getCell(range.startRow, c).value ?? fallback)
          : fallback;
    out.push({ col: c, label: label.length > 0 ? label : fallback });
  }
  return out;
}

function aggregateOptions(): readonly { value: PivotAggregateKind; label: string }[] {
  return [
    { value: "sum", label: "求和" },
    { value: "count", label: "计数" },
    { value: "average", label: "平均值" },
    { value: "max", label: "最大值" },
    { value: "min", label: "最小值" },
  ];
}

export function closePivotTableFieldsPane(): void {
  openPivotFieldsPaneDefId = null;
  if (paneKeydownHandler !== null) {
    document.removeEventListener("keydown", paneKeydownHandler, true);
    paneKeydownHandler = null;
  }
  if (activePaneRoot !== null) {
    activePaneRoot.remove();
    activePaneRoot = null;
  }
}

/**
 * 在选区落在透视输出区域内时，打开右侧「数据透视表字段」窗格（类 Excel）。
 */
export function showPivotTableFieldsPane(
  host: PivotFieldsPaneHost,
  pivotSheet: Worksheet,
  pivotDefId: string,
): void {
  const wb = host.workbook;
  if (wb === undefined) {
    return;
  }
  const def = pivotSheet.getPivotTableDefinitionsSnapshot().find((d) => d.id === pivotDefId);
  if (def === undefined) {
    return;
  }
  const sourceSheet = wb.getSheet(def.sourceSheetIndex);
  if (sourceSheet === undefined) {
    return;
  }

  closePivotTableFieldsPane();
  ensurePaneStyles();

  const n = normalizeSelectionRange(def.sourceRange);
  const allFields = collectSourceFields(sourceSheet, n, def.hasHeaders);
  const labelByCol = new Map(allFields.map((f) => [f.col, f.label]));

  let filterCols = [...def.filterFieldCols];
  let columnCols = [...def.columnFieldCols];
  let rowCols = [...def.rowFieldCols];
  let valueEntries: ValueEntry[] = def.valueFields.map((v, i) => ({
    id: `init-${i}`,
    col: v.col,
    aggregate: v.aggregate,
    computed: v.computed,
  }));
  let valueAxisOnRows = def.valueFieldsOnRows === true;

  enforceExclusivePivotColumns(filterCols, columnCols, rowCols, valueEntries);

  let searchQuery = "";
  let warnText = "";
  let applyTimer: ReturnType<typeof setTimeout> | null = null;
  let dragState: PivotDragState | null = null;
  let pivotDragSourceEl: HTMLElement | null = null;

  const root = document.createElement("aside");
  root.className = "fs-pivot-fields";
  root.setAttribute("role", "complementary");
  root.setAttribute("aria-label", "数据透视表字段");

  const head = document.createElement("div");
  head.className = "fs-pivot-fields__head";
  const title = document.createElement("div");
  title.className = "fs-pivot-fields__title";
  title.textContent = "数据透视表字段";
  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.className = "fs-pivot-fields__close";
  closeBtn.setAttribute("aria-label", "关闭");
  closeBtn.textContent = "×";
  closeBtn.addEventListener("click", () => {
    closePivotTableFieldsPane();
  });
  head.appendChild(title);
  head.appendChild(closeBtn);

  const searchWrap = document.createElement("div");
  searchWrap.className = "fs-pivot-fields__search";
  const searchInner = document.createElement("div");
  searchInner.className = "fs-pivot-fields__search-inner";
  const searchIcon = document.createElement("span");
  searchIcon.className = "fs-pivot-fields__search-icon";
  searchIcon.setAttribute("aria-hidden", "true");
  searchIcon.innerHTML =
    '<svg width="14" height="14" viewBox="0 0 16 16" fill="none" xmlns="http://www.w3.org/2000/svg"><path d="M10.5 9.2h-.55l-.2-.2A4.1 4.1 0 1 0 9.4 10l.2.2v.55l3.2 3.2 1-1-3.2-3.2zm-3.6 0a2.85 2.85 0 1 1 0-5.7 2.85 2.85 0 0 1 0 5.7z" fill="currentColor"/></svg>';
  const searchInput = document.createElement("input");
  searchInput.type = "search";
  searchInput.placeholder = "搜索字段";
  searchInput.setAttribute("autocomplete", "off");
  searchInner.appendChild(searchIcon);
  searchInner.appendChild(searchInput);
  searchWrap.appendChild(searchInner);

  const listWrap = document.createElement("div");
  listWrap.className = "fs-pivot-fields__list-wrap";

  const warn = document.createElement("p");
  warn.className = "fs-pivot-fields__warn";

  const zones = document.createElement("div");
  zones.className = "fs-pivot-fields__zones";

  const mkZone = (
    zoneId: PivotZoneId,
    titleText: string,
    iconMarkup: string,
  ): { zone: HTMLElement; chips: HTMLElement } => {
    const zone = document.createElement("div");
    zone.className = "fs-pivot-fields__zone";
    zone.dataset.zone = zoneId;
    const zh = document.createElement("div");
    zh.className = "fs-pivot-fields__zone-head";
    const zi = document.createElement("div");
    zi.className = "fs-pivot-fields__zone-icon";
    zi.innerHTML = iconMarkup;
    const zt = document.createElement("div");
    zt.className = "fs-pivot-fields__zone-title";
    zt.textContent = titleText;
    zh.appendChild(zi);
    zh.appendChild(zt);
    const chips = document.createElement("div");
    chips.className = "fs-pivot-fields__chips";
    chips.dataset.pivotChips = zoneId;
    zone.appendChild(zh);
    zone.appendChild(chips);
    return { zone, chips };
  };

  const iconFilter =
    '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M2 2.5h12v1L9 8.2V13.5H7V8.2L2 3.5v-1z"/></svg>';
  const iconCols =
    '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M3 3h2.5v10H3V3zm3.75 2h2.5v8h-2.5V5zm3.75-2H13v10h-2.5V3z"/></svg>';
  const iconRows =
    '<svg viewBox="0 0 16 16" aria-hidden="true"><path d="M3 3h10v2.5H3V3zm0 3.75h10v2.5H3v-2.5zm0 3.75h10V13H3v-2.5z"/></svg>';
  const iconSigma =
    '<svg viewBox="0 0 16 16" aria-hidden="true"><text x="1" y="12" font-size="11" font-family="Segoe UI,system-ui,sans-serif" fill="currentColor">Σ</text></svg>';

  const zFilter = mkZone("filter", "筛选器", iconFilter);
  const zCol = mkZone("columns", "列", iconCols);
  const zRow = mkZone("rows", "行", iconRows);
  const zVal = mkZone("values", "值", iconSigma);

  zones.appendChild(zFilter.zone);
  zones.appendChild(zCol.zone);
  zones.appendChild(zRow.zone);
  zones.appendChild(zVal.zone);

  const zonesHint = document.createElement("div");
  zonesHint.className = "fs-pivot-fields__zones-hint";
  zonesHint.textContent = "在区域之间拖动字段";

  root.appendChild(head);
  root.appendChild(searchWrap);
  root.appendChild(listWrap);
  root.appendChild(warn);
  root.appendChild(zones);
  root.appendChild(zonesHint);

  const sheetHost = host.getSheetContainerElement();
  const mountParent = sheetHost.parentElement;
  if (mountParent !== null) {
    root.classList.add("fs-pivot-fields--docked");
    mountParent.insertBefore(root, sheetHost.nextSibling);
  } else {
    root.classList.add("fs-pivot-fields--floating");
    document.body.appendChild(root);
  }
  activePaneRoot = root;

  const isColInAxisZones = (col: number): boolean =>
    filterCols.includes(col) || columnCols.includes(col) || rowCols.includes(col);

  const isFieldChecked = (col: number): boolean =>
    isColInAxisZones(col) || valueEntries.some((e) => e.col === col);

  const removeColFromAxis = (col: number): void => {
    filterCols = filterCols.filter((c) => c !== col);
    columnCols = columnCols.filter((c) => c !== col);
    rowCols = rowCols.filter((c) => c !== col);
  };

  const getAxisArray = (z: "filter" | "columns" | "rows"): number[] => {
    if (z === "filter") {
      return filterCols;
    }
    if (z === "columns") {
      return columnCols;
    }
    return rowCols;
  };

  const allZoneElements = [zFilter.zone, zCol.zone, zRow.zone, zVal.zone];

  const clearDropHighlight = (): void => {
    for (const z of allZoneElements) {
      z.classList.remove("fs-pivot-fields__zone--drop-target");
    }
  };

  const endPivotDrag = (): void => {
    if (pivotDragSourceEl !== null) {
      pivotDragSourceEl.removeAttribute("data-pivot-drag");
    }
    pivotDragSourceEl = null;
    dragState = null;
    clearDropHighlight();
  };

  root.addEventListener("dragend", endPivotDrag, true);

  const applyPivotDrop = (targetZone: PivotZoneId, insertIndex: number): void => {
    const st = dragState;
    if (st === null) {
      return;
    }
    const clamp = (n: number, max: number): number => Math.max(0, Math.min(n, max));

    if (st.kind === "valueAxisChip") {
      if (targetZone === "rows") {
        valueAxisOnRows = true;
      } else if (targetZone === "columns") {
        valueAxisOnRows = false;
      }
      return;
    }

    if (st.kind === "fieldList" && st.col !== undefined) {
      const col = st.col;
      removeColFromAxis(col);
      valueEntries = valueEntries.filter((e) => e.col !== col);
      if (targetZone === "filter") {
        filterCols.splice(clamp(insertIndex, filterCols.length), 0, col);
      } else if (targetZone === "columns") {
        columnCols.splice(clamp(insertIndex, columnCols.length), 0, col);
      } else if (targetZone === "rows") {
        rowCols.splice(clamp(insertIndex, rowCols.length), 0, col);
      } else {
        valueEntries.splice(clamp(insertIndex, valueEntries.length), 0, {
          id: newId(),
          col,
          aggregate: "sum",
          computed: undefined,
        });
      }
      return;
    }

    if (st.kind === "axisChip" && st.col !== undefined && st.fromZone !== undefined) {
      const col = st.col;
      const fromZ = st.fromZone;
      const fromI = st.fromIndex;
      if (fromZ === "values") {
        return;
      }
      if (fromZ === targetZone) {
        const arr = getAxisArray(fromZ);
        arrayMove(arr, fromI, insertIndex);
      } else {
        const fromArr = getAxisArray(fromZ);
        if (fromI < 0 || fromI >= fromArr.length) {
          return;
        }
        fromArr.splice(fromI, 1);
        valueEntries = valueEntries.filter((e) => e.col !== col);
        if (targetZone === "values") {
          valueEntries.splice(clamp(insertIndex, valueEntries.length), 0, {
            id: newId(),
            col,
            aggregate: "sum",
            computed: undefined,
          });
        } else {
          const toArr = getAxisArray(targetZone);
          toArr.splice(clamp(insertIndex, toArr.length), 0, col);
        }
      }
      return;
    }

    if (st.kind === "valueChip") {
      const fromI = st.fromIndex;
      if (fromI < 0 || fromI >= valueEntries.length) {
        return;
      }
      if (targetZone === "values") {
        arrayMove(valueEntries, fromI, insertIndex);
      } else {
        const entry = valueEntries[fromI];
        valueEntries.splice(fromI, 1);
        removeColFromAxis(entry.col);
        const toArr = getAxisArray(targetZone);
        toArr.splice(clamp(insertIndex, toArr.length), 0, entry.col);
      }
    }
  };

  const scheduleApply = (): void => {
    if (applyTimer !== null) {
      clearTimeout(applyTimer);
    }
    applyTimer = setTimeout(() => {
      applyTimer = null;
      void applyLayout();
    }, 280);
  };

  const applyLayout = (): void => {
    warnText = "";
    if (rowCols.length === 0 || valueEntries.length === 0) {
      warnText = "至少需要 1 个行字段和 1 个值字段。";
      warn.textContent = warnText;
      return;
    }
    if (columnCols.length > 0 && valueEntries.length > 1) {
      warnText = "已放置列字段时，仅支持单个值字段（与 Excel 行为一致）。";
      warn.textContent = warnText;
      return;
    }
    warn.textContent = "";
    const cmd = new UpdatePivotTableLayoutCommand(wb, pivotSheet, pivotDefId, {
      sourceRange: { ...def.sourceRange },
      hasHeaders: def.hasHeaders,
      rowFieldCols: rowCols,
      columnFieldCols: columnCols,
      filterFieldCols: filterCols,
      valueFields: valueEntries.map((e) => ({
        col: e.col,
        aggregate: e.aggregate,
        ...(e.computed !== undefined ? { computed: e.computed } : {}),
      })),
      valueFieldsOnRows:
        columnCols.length > 0 || valueEntries.length <= 1 ? false : valueAxisOnRows,
    });
    host.workspace.commands.execute(cmd);
    host.refresh();
  };

  const renderFieldList = (): void => {
    listWrap.replaceChildren();
    const q = searchQuery.trim().toLowerCase();
    for (const f of allFields) {
      if (q.length > 0 && !f.label.toLowerCase().includes(q)) {
        continue;
      }
      const row = document.createElement("div");
      row.className = "fs-pivot-fields__row";
      row.draggable = true;
      row.dataset.col = String(f.col);
      const cb = document.createElement("input");
      cb.type = "checkbox";
      cb.checked = isFieldChecked(f.col);
      cb.addEventListener("change", () => {
        if (cb.checked) {
          valueEntries = valueEntries.filter((e) => e.col !== f.col);
          removeColFromAxis(f.col);
          if (!rowCols.includes(f.col)) {
            rowCols.push(f.col);
          }
        } else {
          removeColFromAxis(f.col);
          valueEntries = valueEntries.filter((e) => e.col !== f.col);
        }
        renderAll();
        scheduleApply();
      });
      const lab = document.createElement("span");
      lab.textContent = f.label;
      lab.style.flex = "1";
      row.appendChild(cb);
      row.appendChild(lab);
      row.addEventListener("dragstart", (ev) => {
        dragState = { kind: "fieldList", col: f.col, fromIndex: 0 };
        pivotDragSourceEl = row;
        row.setAttribute("data-pivot-drag", "1");
        ev.dataTransfer?.setData("text/plain", String(f.col));
        ev.dataTransfer!.effectAllowed = "move";
      });
      listWrap.appendChild(row);
    }
  };

  const makeAxisChip = (
    col: number,
    zone: "filter" | "columns" | "rows",
    indexInZone: number,
    onRemove: () => void,
  ): HTMLElement => {
    const chip = document.createElement("div");
    chip.className = "fs-pivot-fields__chip";
    chip.draggable = true;
    chip.dataset.col = String(col);
    const handle = document.createElement("div");
    handle.className = "fs-pivot-fields__chip-handle";
    handle.innerHTML = "<span></span>";
    const body = document.createElement("div");
    body.className = "fs-pivot-fields__chip-body";
    const text = document.createElement("span");
    text.className = "fs-pivot-fields__chip-label";
    const axisLabel = labelByCol.get(col) ?? `列${col}`;
    text.textContent = axisLabel;
    text.title = axisLabel;
    const info = document.createElement("button");
    info.type = "button";
    info.className = "fs-pivot-fields__chip-info";
    info.textContent = "i";
    info.title = labelByCol.get(col) ?? `列${col}`;
    info.addEventListener("click", (e) => {
      e.stopPropagation();
    });
    const xb = document.createElement("button");
    xb.type = "button";
    xb.className = "fs-pivot-fields__chip-x";
    xb.setAttribute("aria-label", "移除");
    xb.textContent = "×";
    xb.addEventListener("click", (e) => {
      e.stopPropagation();
      onRemove();
      renderAll();
      scheduleApply();
    });
    body.appendChild(text);
    body.appendChild(info);
    body.appendChild(xb);
    chip.appendChild(handle);
    chip.appendChild(body);
    chip.addEventListener("dragstart", (ev) => {
      dragState = { kind: "axisChip", col, fromZone: zone, fromIndex: indexInZone };
      pivotDragSourceEl = chip;
      chip.setAttribute("data-pivot-drag", "1");
      ev.dataTransfer?.setData("text/plain", String(col));
      ev.dataTransfer!.effectAllowed = "move";
    });
    return chip;
  };

  const makeValueAxisChip = (zone: "rows" | "columns"): HTMLElement => {
    const axisChip = document.createElement("div");
    axisChip.className = "fs-pivot-fields__chip fs-pivot-fields__chip--value-axis";
    axisChip.draggable = true;
    axisChip.dataset.pivotValueAxis = "1";
    const axHandle = document.createElement("div");
    axHandle.className = "fs-pivot-fields__chip-handle";
    axHandle.innerHTML = "<span></span>";
    const axBody = document.createElement("div");
    axBody.className = "fs-pivot-fields__chip-body";
    const axLab = document.createElement("span");
    axLab.className = "fs-pivot-fields__chip-label";
    axLab.textContent = "数值";
    axLab.title = "在行或列区域间拖动，切换度量在行或列上展开（与 Excel「数值」一致）";
    axBody.appendChild(axLab);
    axisChip.appendChild(axHandle);
    axisChip.appendChild(axBody);
    axisChip.addEventListener("dragstart", (ev) => {
      dragState = { kind: "valueAxisChip", fromZone: zone, fromIndex: 0 };
      pivotDragSourceEl = axisChip;
      axisChip.setAttribute("data-pivot-drag", "1");
      ev.dataTransfer?.setData("text/plain", "value-axis");
      ev.dataTransfer!.effectAllowed = "move";
    });
    return axisChip;
  };

  const renderZones = (): void => {
    zFilter.chips.replaceChildren();
    zCol.chips.replaceChildren();
    zRow.chips.replaceChildren();
    zVal.chips.replaceChildren();

    for (let i = 0; i < filterCols.length; i++) {
      const c = filterCols[i];
      zFilter.chips.appendChild(
        makeAxisChip(c, "filter", i, () => {
          filterCols = filterCols.filter((x) => x !== c);
        }),
      );
    }
    for (let i = 0; i < columnCols.length; i++) {
      const c = columnCols[i];
      zCol.chips.appendChild(
        makeAxisChip(c, "columns", i, () => {
          columnCols = columnCols.filter((x) => x !== c);
        }),
      );
    }
    for (let i = 0; i < rowCols.length; i++) {
      const c = rowCols[i];
      zRow.chips.appendChild(
        makeAxisChip(c, "rows", i, () => {
          rowCols = rowCols.filter((x) => x !== c);
        }),
      );
    }
    if (valueEntries.length > 1 && columnCols.length === 0 && valueAxisOnRows) {
      zRow.chips.appendChild(makeValueAxisChip("rows"));
    }
    if (valueEntries.length > 1 && columnCols.length === 0 && !valueAxisOnRows) {
      zCol.chips.appendChild(makeValueAxisChip("columns"));
    }
    for (let vi = 0; vi < valueEntries.length; vi++) {
      const e = valueEntries[vi];
      const chip = document.createElement("div");
      chip.className = "fs-pivot-fields__chip fs-pivot-fields__chip--values";
      chip.draggable = true;
      chip.dataset.entryId = e.id;
      const handle = document.createElement("div");
      handle.className = "fs-pivot-fields__chip-handle";
      handle.innerHTML = "<span></span>";
      const body = document.createElement("div");
      body.className = "fs-pivot-fields__chip-body";
      const nm = labelByCol.get(e.col) ?? "值";
      const text = document.createElement("span");
      text.className = "fs-pivot-fields__chip-label";
      const denNm =
        e.computed?.kind === "bucketRatio"
          ? (labelByCol.get(e.computed.denominatorCol) ?? "")
          : undefined;
      const valueCaption = getPivotValueFieldCaption(
        { col: e.col, aggregate: e.aggregate, computed: e.computed },
        nm,
        denNm,
      );
      text.textContent = valueCaption;
      text.title = valueCaption;
      const info = document.createElement("button");
      info.type = "button";
      info.className = "fs-pivot-fields__chip-info";
      info.textContent = "i";
      info.title = "汇总方式";
      info.addEventListener("click", (ev) => {
        ev.stopPropagation();
        const menu = document.createElement("div");
        menu.style.cssText =
          "position:fixed;z-index:13000;background:#fff;border:1px solid #8a8886;border-radius:4px;box-shadow:0 2px 8px rgba(0,0,0,.15);padding:4px 0;min-width:128px;";
        const rect = info.getBoundingClientRect();
        menu.style.left = `${Math.min(rect.left, window.innerWidth - 140)}px`;
        menu.style.top = `${rect.bottom + 4}px`;
        const dismiss = (): void => {
          menu.remove();
          document.removeEventListener("click", dismiss, true);
        };
        for (const o of aggregateOptions()) {
          const b = document.createElement("button");
          b.type = "button";
          b.textContent = o.label;
          b.style.cssText =
            "display:block;width:100%;text-align:left;padding:8px 12px;border:none;background:transparent;cursor:pointer;font-size:12px;color:#323130;";
          const isActive = e.computed === undefined && o.value === e.aggregate;
          if (isActive) {
            b.style.fontWeight = "600";
            b.style.background = "#f3f2f1";
          }
          b.addEventListener("click", (evt) => {
            evt.stopPropagation();
            e.aggregate = o.value;
            e.computed = undefined;
            dismiss();
            renderAll();
            scheduleApply();
          });
          b.addEventListener("mouseenter", () => {
            b.style.background = "#edebe9";
          });
          b.addEventListener("mouseleave", () => {
            b.style.background = isActive ? "#f3f2f1" : "transparent";
          });
          menu.appendChild(b);
        }
        const sep = document.createElement("div");
        sep.style.cssText = "height:1px;background:#edebe9;margin:4px 0;";
        menu.appendChild(sep);
        const shareBtn = document.createElement("button");
        shareBtn.type = "button";
        shareBtn.textContent = "占比（÷ 数据源该列总计）";
        shareBtn.style.cssText =
          "display:block;width:100%;text-align:left;padding:8px 12px;border:none;background:transparent;cursor:pointer;font-size:12px;color:#323130;";
        if (e.computed?.kind === "shareOfGrandTotal") {
          shareBtn.style.fontWeight = "600";
          shareBtn.style.background = "#f3f2f1";
        }
        shareBtn.addEventListener("click", (evt) => {
          evt.stopPropagation();
          e.aggregate = "sum";
          e.computed = { kind: "shareOfGrandTotal" };
          dismiss();
          renderAll();
          scheduleApply();
        });
        shareBtn.addEventListener("mouseenter", () => {
          shareBtn.style.background = "#edebe9";
        });
        shareBtn.addEventListener("mouseleave", () => {
          shareBtn.style.background =
            e.computed?.kind === "shareOfGrandTotal" ? "#f3f2f1" : "transparent";
        });
        menu.appendChild(shareBtn);
        const ratioHint = document.createElement("div");
        ratioHint.textContent = "比率（分组内求和后再相除）";
        ratioHint.style.cssText = "padding:4px 12px 2px;font-size:11px;color:#605e5c;";
        menu.appendChild(ratioHint);
        for (const f of allFields) {
          if (f.col === e.col) {
            continue;
          }
          const rb = document.createElement("button");
          rb.type = "button";
          rb.textContent = `÷ ${f.label}`;
          rb.style.cssText =
            "display:block;width:100%;text-align:left;padding:8px 12px;border:none;background:transparent;cursor:pointer;font-size:12px;color:#323130;";
          const ratioOn = e.computed?.kind === "bucketRatio" && e.computed.denominatorCol === f.col;
          if (ratioOn) {
            rb.style.fontWeight = "600";
            rb.style.background = "#f3f2f1";
          }
          rb.addEventListener("click", (evt) => {
            evt.stopPropagation();
            e.aggregate = "sum";
            e.computed = { kind: "bucketRatio", denominatorCol: f.col };
            dismiss();
            renderAll();
            scheduleApply();
          });
          rb.addEventListener("mouseenter", () => {
            rb.style.background = "#edebe9";
          });
          rb.addEventListener("mouseleave", () => {
            rb.style.background = ratioOn ? "#f3f2f1" : "transparent";
          });
          menu.appendChild(rb);
        }
        document.body.appendChild(menu);
        queueMicrotask(() => document.addEventListener("click", dismiss, true));
      });
      const xb = document.createElement("button");
      xb.type = "button";
      xb.className = "fs-pivot-fields__chip-x";
      xb.setAttribute("aria-label", "移除");
      xb.textContent = "×";
      xb.addEventListener("click", (evt) => {
        evt.stopPropagation();
        valueEntries = valueEntries.filter((x) => x.id !== e.id);
        renderAll();
        scheduleApply();
      });
      body.appendChild(text);
      body.appendChild(info);
      body.appendChild(xb);
      chip.appendChild(handle);
      chip.appendChild(body);
      chip.addEventListener("dragstart", (ev) => {
        dragState = { kind: "valueChip", entryId: e.id, fromZone: "values", fromIndex: vi };
        pivotDragSourceEl = chip;
        chip.setAttribute("data-pivot-drag", "1");
        ev.dataTransfer?.setData("text/plain", e.id);
        ev.dataTransfer!.effectAllowed = "move";
      });
      zVal.chips.appendChild(chip);
    }
  };

  const renderAll = (): void => {
    enforceExclusivePivotColumns(filterCols, columnCols, rowCols, valueEntries);
    warn.textContent = warnText;
    renderFieldList();
    renderZones();
  };

  const bindZoneDrop = (zoneEl: HTMLElement, chipsEl: HTMLElement, zone: PivotZoneId): void => {
    zoneEl.addEventListener("dragover", (ev) => {
      if (dragState === null) {
        return;
      }
      ev.preventDefault();
      ev.dataTransfer!.dropEffect = "move";
      clearDropHighlight();
      zoneEl.classList.add("fs-pivot-fields__zone--drop-target");
    });
    zoneEl.addEventListener("drop", (ev) => {
      if (dragState === null) {
        return;
      }
      ev.preventDefault();
      clearDropHighlight();
      const insertIdx = insertIndexFromPointer(chipsEl, ev.clientY);
      applyPivotDrop(zone, insertIdx);
      renderAll();
      scheduleApply();
    });
  };

  bindZoneDrop(zFilter.zone, zFilter.chips, "filter");
  bindZoneDrop(zCol.zone, zCol.chips, "columns");
  bindZoneDrop(zRow.zone, zRow.chips, "rows");
  bindZoneDrop(zVal.zone, zVal.chips, "values");

  searchInput.addEventListener("input", () => {
    searchQuery = searchInput.value;
    renderFieldList();
  });

  renderAll();

  const onKey = (ev: KeyboardEvent): void => {
    if (ev.key === "Escape") {
      ev.preventDefault();
      closePivotTableFieldsPane();
    }
  };
  paneKeydownHandler = onKey;
  document.addEventListener("keydown", onKey, true);

  openPivotFieldsPaneDefId = pivotDefId;
}

/**
 * 按活动单元格同步字段窗格：在透视输出区域内则显示（同表不重复重建），否则关闭。
 */
export function syncPivotTableFieldsPaneWithSelection(
  host: PivotFieldsPaneHost,
  sheet: Worksheet | undefined,
  activeRow: number,
  activeCol: number,
): void {
  if (sheet === undefined) {
    closePivotTableFieldsPane();
    return;
  }
  const def = findPivotTableDefinitionAtCell(sheet, activeRow, activeCol);
  if (def === null) {
    closePivotTableFieldsPane();
    return;
  }
  if (openPivotFieldsPaneDefId === def.id && activePaneRoot !== null) {
    return;
  }
  showPivotTableFieldsPane(host, sheet, def.id);
}

/** 若活动单元格落在透视区域内则打开字段窗格，否则关闭。 */
export function tryOpenPivotFieldsPaneForSelection(
  host: PivotFieldsPaneHost,
  sheet: Worksheet,
  activeRow: number,
  activeCol: number,
): boolean {
  const def = findPivotTableDefinitionAtCell(sheet, activeRow, activeCol);
  syncPivotTableFieldsPaneWithSelection(host, sheet, activeRow, activeCol);
  return def !== null;
}
