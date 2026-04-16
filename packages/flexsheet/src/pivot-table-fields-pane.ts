import {
  normalizeSelectionRange,
  type ICommand,
  type PivotAggregateKind,
  type Workbook,
  type Worksheet,
} from "@flexsheet/core";
import {
  dataFieldCaption,
  findPivotTableDefinitionAtCell,
  UpdatePivotTableLayoutCommand,
} from "./pivot-table-command.js";

export interface PivotFieldsPaneHost {
  readonly workbook: Workbook | undefined;
  readonly workspace: { readonly commands: { execute(cmd: ICommand): void } };
  refresh(): void;
}

interface SourceField {
  readonly col: number;
  readonly label: string;
}

interface ValueEntry {
  readonly id: string;
  col: number;
  aggregate: PivotAggregateKind;
}

let paneStylesInjected = false;
let activePaneRoot: HTMLElement | null = null;
let paneKeydownHandler: ((ev: KeyboardEvent) => void) | null = null;

function ensurePaneStyles(): void {
  if (paneStylesInjected) {
    return;
  }
  paneStylesInjected = true;
  const style = document.createElement("style");
  style.setAttribute("data-fs-pivot-fields-pane", "1");
  style.textContent = `
.fs-pivot-fields {
  position: fixed;
  z-index: 12000;
  top: 72px;
  right: 12px;
  bottom: 12px;
  width: 300px;
  max-width: calc(100vw - 24px);
  background: #f3f2f1;
  border: 1px solid #d2d0ce;
  border-radius: 4px;
  box-shadow: 0 4px 16px rgba(0,0,0,.12);
  display: flex;
  flex-direction: column;
  font-family: "Segoe UI", system-ui, -apple-system, sans-serif;
  font-size: 12px;
  color: #323130;
}
.fs-pivot-fields__head {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 10px 12px;
  border-bottom: 1px solid #edebe9;
  background: #faf9f8;
  flex-shrink: 0;
}
.fs-pivot-fields__title {
  font-weight: 600;
  font-size: 13px;
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
  padding: 8px 12px;
  flex-shrink: 0;
}
.fs-pivot-fields__search input {
  width: 100%;
  box-sizing: border-box;
  padding: 6px 8px;
  border: 1px solid #8a8886;
  border-radius: 2px;
  font-size: 12px;
}
.fs-pivot-fields__list-wrap {
  flex: 0 1 38%;
  min-height: 120px;
  overflow: auto;
  margin: 0 8px 8px;
  background: #fff;
  border: 1px solid #edebe9;
  border-radius: 2px;
}
.fs-pivot-fields__row {
  display: flex;
  align-items: center;
  gap: 8px;
  padding: 6px 10px;
  border-bottom: 1px solid #f3f2f1;
  cursor: grab;
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
  padding: 0 8px 10px;
  display: flex;
  flex-direction: column;
  gap: 8px;
}
.fs-pivot-fields__zone {
  background: #fff;
  border: 1px solid #edebe9;
  border-radius: 2px;
  padding: 8px;
  min-height: 56px;
}
.fs-pivot-fields__zone-title {
  font-size: 11px;
  color: #605e5c;
  margin-bottom: 6px;
  font-weight: 600;
}
.fs-pivot-fields__chips {
  display: flex;
  flex-wrap: wrap;
  gap: 6px;
  min-height: 28px;
}
.fs-pivot-fields__chip {
  display: inline-flex;
  align-items: center;
  gap: 4px;
  padding: 4px 8px;
  background: #e1f0fe;
  border: 1px solid #8ec5f6;
  border-radius: 4px;
  font-size: 11px;
  max-width: 100%;
}
.fs-pivot-fields__chip--values {
  background: #fff4ce;
  border-color: #f5d142;
}
.fs-pivot-fields__chip-x {
  border: none;
  background: transparent;
  cursor: pointer;
  padding: 0 2px;
  font-size: 14px;
  line-height: 1;
  color: #605e5c;
}
.fs-pivot-fields__chip-x:hover {
  color: #a4262c;
}
.fs-pivot-fields__chip-info {
  border: none;
  background: transparent;
  cursor: pointer;
  font-size: 11px;
  font-weight: 700;
  font-style: italic;
  color: #0078d4;
  padding: 0 2px;
}
.fs-pivot-fields__warn {
  margin: 0 8px 8px;
  font-size: 11px;
  color: #a4262c;
  min-height: 14px;
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
  }));

  let searchQuery = "";
  let warnText = "";
  let applyTimer: ReturnType<typeof setTimeout> | null = null;

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
  const searchInput = document.createElement("input");
  searchInput.type = "search";
  searchInput.placeholder = "搜索字段";
  searchInput.setAttribute("autocomplete", "off");
  searchWrap.appendChild(searchInput);

  const listWrap = document.createElement("div");
  listWrap.className = "fs-pivot-fields__list-wrap";

  const warn = document.createElement("p");
  warn.className = "fs-pivot-fields__warn";

  const zones = document.createElement("div");
  zones.className = "fs-pivot-fields__zones";

  const mkZone = (zoneId: string, titleText: string): { zone: HTMLElement; chips: HTMLElement } => {
    const zone = document.createElement("div");
    zone.className = "fs-pivot-fields__zone";
    zone.dataset.zone = zoneId;
    const zt = document.createElement("div");
    zt.className = "fs-pivot-fields__zone-title";
    zt.textContent = titleText;
    const chips = document.createElement("div");
    chips.className = "fs-pivot-fields__chips";
    zone.appendChild(zt);
    zone.appendChild(chips);
    return { zone, chips };
  };

  const zFilter = mkZone("filter", "筛选器");
  const zCol = mkZone("columns", "列");
  const zRow = mkZone("rows", "行");
  const zVal = mkZone("values", "∑ 值");

  zones.appendChild(zFilter.zone);
  zones.appendChild(zCol.zone);
  zones.appendChild(zRow.zone);
  zones.appendChild(zVal.zone);

  root.appendChild(head);
  root.appendChild(searchWrap);
  root.appendChild(listWrap);
  root.appendChild(warn);
  root.appendChild(zones);
  document.body.appendChild(root);
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
      valueFields: valueEntries.map((e) => ({ col: e.col, aggregate: e.aggregate })),
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
          if (!isColInAxisZones(f.col) && !valueEntries.some((e) => e.col === f.col)) {
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
        ev.dataTransfer?.setData("text/fs-pivot-col", String(f.col));
        ev.dataTransfer?.setData("text/plain", String(f.col));
        ev.dataTransfer!.effectAllowed = "move";
      });
      listWrap.appendChild(row);
    }
  };

  const makeChip = (
    col: number,
    zone: "filter" | "columns" | "rows",
    onRemove: () => void,
  ): HTMLElement => {
    const chip = document.createElement("div");
    chip.className = "fs-pivot-fields__chip";
    chip.draggable = true;
    chip.dataset.col = String(col);
    const text = document.createElement("span");
    text.textContent = labelByCol.get(col) ?? `列${col}`;
    text.style.overflow = "hidden";
    text.style.textOverflow = "ellipsis";
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
    chip.appendChild(text);
    chip.appendChild(xb);
    chip.addEventListener("dragstart", (ev) => {
      ev.dataTransfer?.setData("text/fs-pivot-zone", zone);
      ev.dataTransfer?.setData("text/fs-pivot-col", String(col));
      ev.dataTransfer!.effectAllowed = "move";
    });
    return chip;
  };

  const renderZones = (): void => {
    zFilter.chips.replaceChildren();
    zCol.chips.replaceChildren();
    zRow.chips.replaceChildren();
    zVal.chips.replaceChildren();

    for (const c of filterCols) {
      zFilter.chips.appendChild(
        makeChip(c, "filter", () => {
          filterCols = filterCols.filter((x) => x !== c);
        }),
      );
    }
    for (const c of columnCols) {
      zCol.chips.appendChild(
        makeChip(c, "columns", () => {
          columnCols = columnCols.filter((x) => x !== c);
        }),
      );
    }
    for (const c of rowCols) {
      zRow.chips.appendChild(
        makeChip(c, "rows", () => {
          rowCols = rowCols.filter((x) => x !== c);
        }),
      );
    }
    for (const e of valueEntries) {
      const chip = document.createElement("div");
      chip.className = "fs-pivot-fields__chip fs-pivot-fields__chip--values";
      chip.draggable = true;
      chip.dataset.entryId = e.id;
      const nm = labelByCol.get(e.col) ?? "值";
      const text = document.createElement("span");
      text.textContent = dataFieldCaption(e.aggregate, nm);
      text.style.overflow = "hidden";
      text.style.textOverflow = "ellipsis";
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
          if (o.value === e.aggregate) {
            b.style.fontWeight = "600";
            b.style.background = "#f3f2f1";
          }
          b.addEventListener("click", (evt) => {
            evt.stopPropagation();
            e.aggregate = o.value;
            dismiss();
            renderAll();
            scheduleApply();
          });
          b.addEventListener("mouseenter", () => {
            b.style.background = "#edebe9";
          });
          b.addEventListener("mouseleave", () => {
            b.style.background = o.value === e.aggregate ? "#f3f2f1" : "transparent";
          });
          menu.appendChild(b);
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
      chip.appendChild(text);
      chip.appendChild(info);
      chip.appendChild(xb);
      zVal.chips.appendChild(chip);
    }
  };

  const bindDrop = (target: HTMLElement, zone: "filter" | "columns" | "rows" | "values"): void => {
    target.addEventListener("dragover", (ev) => {
      ev.preventDefault();
      ev.dataTransfer!.dropEffect = "move";
    });
    target.addEventListener("drop", (ev) => {
      ev.preventDefault();
      const raw = ev.dataTransfer?.getData("text/fs-pivot-col") ?? "";
      const col = Number.parseInt(raw, 10);
      if (!Number.isInteger(col)) {
        return;
      }
      removeColFromAxis(col);
      valueEntries = valueEntries.filter((e) => e.col !== col);
      if (zone === "filter") {
        if (!filterCols.includes(col)) {
          filterCols.push(col);
        }
      } else if (zone === "columns") {
        if (!columnCols.includes(col)) {
          columnCols.push(col);
        }
      } else if (zone === "rows") {
        if (!rowCols.includes(col)) {
          rowCols.push(col);
        }
      } else {
        valueEntries.push({ id: newId(), col, aggregate: "sum" });
      }
      renderAll();
      scheduleApply();
    });
  };

  bindDrop(zFilter.zone, "filter");
  bindDrop(zCol.zone, "columns");
  bindDrop(zRow.zone, "rows");
  bindDrop(zVal.zone, "values");

  const renderAll = (): void => {
    warn.textContent = warnText;
    renderFieldList();
    renderZones();
  };

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
}

/** 若活动单元格落在透视区域内则打开字段窗格，否则无操作。 */
export function tryOpenPivotFieldsPaneForSelection(
  host: PivotFieldsPaneHost,
  sheet: Worksheet,
  activeRow: number,
  activeCol: number,
): boolean {
  const def = findPivotTableDefinitionAtCell(sheet, activeRow, activeCol);
  if (def === null) {
    return false;
  }
  showPivotTableFieldsPane(host, sheet, def.id);
  return true;
}
