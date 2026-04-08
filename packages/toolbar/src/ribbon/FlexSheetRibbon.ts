import type { SheetTheme } from "@flexsheet/theme";
import {
  mountDataTab,
  mountFormulaTab,
  mountHomeTab,
  mountInsertTab,
  mountPageLayoutTab,
  mountViewTab,
} from "./tabs/index.js";
import type {
  FlexSheetLike,
  FlexSheetRibbonOptions,
  RibbonCommandEvent,
  RibbonTabId,
} from "./ribbon-types.js";
import type { RibbonEmit } from "../toolbar/toolbar-button.js";
import { closeAllRibbonPopups } from "../toolbar/toolbar-dropdown.js";
import { createRibbonBackstage, type RibbonBackstageHandles } from "./ribbon-backstage.js";

import "./FlexSheetRibbon.css";

const TAB_ORDER: RibbonTabId[] = ["home", "insert", "pageLayout", "formula", "data", "view"];

const TAB_LABEL: Record<RibbonTabId, string> = {
  home: "开始",
  insert: "插入",
  pageLayout: "页面布局",
  formula: "公式",
  data: "数据",
  view: "视图",
};

export class FlexSheetRibbon {
  private readonly root: HTMLElement;
  private readonly coverRoot: HTMLElement;
  private readonly onCommand?: (ev: RibbonCommandEvent) => void;
  private flexSheet: FlexSheetLike | undefined;
  private activeTab: RibbonTabId = "home";
  private readonly tabButtons = new Map<RibbonTabId, HTMLButtonElement>();
  private readonly panels = new Map<RibbonTabId, HTMLElement>();
  private readonly onDocPointerDown: (e: PointerEvent) => void;
  private backstage: RibbonBackstageHandles | null = null;

  constructor(options: FlexSheetRibbonOptions) {
    this.onCommand = options.onCommand;
    this.flexSheet = options.flexSheet;
    this.coverRoot = options.backstageCoverRoot ?? options.container.parentElement ?? document.body;
    this.root = document.createElement("div");
    this.root.className = "fs-ribbon";
    this.root.setAttribute("role", "region");
    this.root.setAttribute("aria-label", "功能区");

    const header = document.createElement("div");
    header.className = "fs-ribbon__header";

    const fileTab = document.createElement("button");
    fileTab.type = "button";
    fileTab.className = "fs-ribbon__tab";
    fileTab.textContent = "文件";
    fileTab.setAttribute("aria-label", "文件菜单");
    fileTab.addEventListener("click", () => {
      this.openBackstage();
    });

    const tablist = document.createElement("div");
    tablist.className = "fs-ribbon__tablist";
    tablist.setAttribute("role", "tablist");
    tablist.setAttribute("aria-label", "主选项卡");
    tablist.appendChild(fileTab);

    for (const id of TAB_ORDER) {
      const btn = document.createElement("button");
      btn.type = "button";
      btn.className = "fs-ribbon__tab";
      btn.setAttribute("role", "tab");
      btn.id = `fs-ribbon-tab-${id}`;
      btn.setAttribute("aria-controls", `fs-ribbon-panel-${id}`);
      btn.setAttribute("tabindex", id === this.activeTab ? "0" : "-1");
      btn.dataset.tabId = id;
      btn.textContent = TAB_LABEL[id];
      btn.addEventListener("click", () => {
        this.selectTab(id);
      });
      btn.addEventListener("keydown", (ev) => {
        this.onTabKeydown(ev, id);
      });
      this.tabButtons.set(id, btn);
      tablist.appendChild(btn);
    }

    header.appendChild(tablist);

    const body = document.createElement("div");
    body.className = "fs-ribbon__body";

    const emit: RibbonEmit = (id, tab, payload) => {
      const ev: RibbonCommandEvent = { id, tab, payload };
      this.onCommand?.(ev);
    };

    for (const id of TAB_ORDER) {
      const panel = document.createElement("div");
      panel.className = "fs-ribbon__panel";
      panel.id = `fs-ribbon-panel-${id}`;
      panel.setAttribute("role", "tabpanel");
      panel.setAttribute("aria-labelledby", `fs-ribbon-tab-${id}`);
      panel.hidden = id !== this.activeTab;
      panel.dataset.tabId = id;
      switch (id) {
        case "home": {
          const homeHandles = mountHomeTab(panel, emit);
          options.onHomeTabMounted?.(homeHandles);
          break;
        }
        case "insert":
          mountInsertTab(panel, emit);
          break;
        case "pageLayout":
          mountPageLayoutTab(panel, emit);
          break;
        case "formula":
          mountFormulaTab(panel, emit);
          break;
        case "data":
          mountDataTab(panel, emit);
          break;
        case "view": {
          const viewHandles = mountViewTab(panel, emit);
          options.onViewTabMounted?.(viewHandles);
          break;
        }
      }
      this.panels.set(id, panel);
      body.appendChild(panel);
    }

    this.root.appendChild(header);
    this.root.appendChild(body);

    options.container.appendChild(this.root);

    this.onDocPointerDown = (e: PointerEvent): void => {
      const t = e.target as HTMLElement | null;
      if (t !== null && t.closest(".fs-dd") !== null) {
        return;
      }
      if (t !== null && t.closest("[data-fs-floating-menu]") !== null) {
        return;
      }
      this.closeAllDropdowns();
    };
    document.addEventListener("pointerdown", this.onDocPointerDown, true);

    if (this.flexSheet !== undefined) {
      this.applyThemeFromSheet(this.flexSheet.getTheme());
    } else {
      this.applyThemeMode("light");
    }
    this.syncTabSelectionUi();
  }

  getElement(): HTMLElement {
    return this.root;
  }

  setFlexSheet(flexSheet: FlexSheetLike | undefined): void {
    this.flexSheet = flexSheet;
    if (flexSheet !== undefined) {
      this.applyThemeFromSheet(flexSheet.getTheme());
    }
  }

  /** 从 SheetTheme 同步 Ribbon 的 CSS 变量（与画布明暗一致；浅色强调色与 Excel 选区绿一致） */
  applyThemeFromSheet(theme: SheetTheme): void {
    const mode = theme.mode;
    this.applyThemeMode(mode);
    this.root.style.setProperty("--fs-ribbon-chrome-text", theme.menuColor);
    this.root.style.setProperty(
      "--fs-ribbon-group-label",
      theme.mode === "dark" ? "#a19f9d" : "#605e5c",
    );
    this.root.style.setProperty("--fs-ribbon-border", theme.menuSeparator);
    this.root.style.setProperty("--fs-ribbon-hover", theme.menuHoverBg);
    this.root.style.setProperty("--fs-ribbon-panel-bg", theme.menuBg);
    if (theme.mode === "light") {
      this.root.style.setProperty("--fs-ribbon-accent", theme.activeCellBorderColor);
      this.root.style.setProperty("--fs-ribbon-accent-press", "#185c37");
    } else {
      this.root.style.removeProperty("--fs-ribbon-accent");
      this.root.style.removeProperty("--fs-ribbon-accent-press");
    }
    this.backstage?.applyThemeMode(theme.mode);
  }

  private applyThemeMode(mode: "light" | "dark"): void {
    this.root.dataset.theme = mode;
  }

  private closeAllDropdowns(): void {
    closeAllRibbonPopups();
  }

  private openBackstage(): void {
    this.closeAllDropdowns();
    if (this.backstage !== null) {
      return;
    }
    const handles = createRibbonBackstage(() => {
      this.closeBackstage();
    }, { flexSheet: this.flexSheet });
    handles.applyThemeMode(this.root.dataset.theme === "dark" ? "dark" : "light");
    this.coverRoot.appendChild(handles.root);
    this.backstage = handles;
    const back = handles.root.querySelector<HTMLButtonElement>(".fs-ribbon-backstage__back");
    back?.focus();
  }

  private closeBackstage(): void {
    if (this.backstage === null) {
      return;
    }
    this.backstage.root.remove();
    this.backstage = null;
  }

  private selectTab(id: RibbonTabId): void {
    this.activeTab = id;
    this.syncTabSelectionUi();
  }

  private syncTabSelectionUi(): void {
    for (const [tid, btn] of this.tabButtons) {
      const selected = tid === this.activeTab;
      btn.setAttribute("aria-selected", selected ? "true" : "false");
      btn.classList.toggle("fs-ribbon__tab--active", selected);
      btn.setAttribute("tabindex", selected ? "0" : "-1");
      const panel = this.panels.get(tid);
      if (panel !== undefined) {
        panel.hidden = !selected;
      }
    }
  }

  private onTabKeydown(ev: KeyboardEvent, id: RibbonTabId): void {
    const idx = TAB_ORDER.indexOf(id);
    if (idx < 0) {
      return;
    }
    let nextIdx = idx;
    if (ev.key === "ArrowRight" || ev.key === "ArrowDown") {
      nextIdx = (idx + 1) % TAB_ORDER.length;
      ev.preventDefault();
    } else if (ev.key === "ArrowLeft" || ev.key === "ArrowUp") {
      nextIdx = (idx - 1 + TAB_ORDER.length) % TAB_ORDER.length;
      ev.preventDefault();
    } else if (ev.key === "Home") {
      nextIdx = 0;
      ev.preventDefault();
    } else if (ev.key === "End") {
      nextIdx = TAB_ORDER.length - 1;
      ev.preventDefault();
    } else {
      return;
    }
    const nextId = TAB_ORDER[nextIdx];
    if (nextId !== undefined) {
      this.selectTab(nextId);
      this.tabButtons.get(nextId)?.focus();
    }
  }

  destroy(): void {
    document.removeEventListener("pointerdown", this.onDocPointerDown, true);
    this.closeBackstage();
    this.root.remove();
  }
}

export type {
  FlexSheetRibbonOptions,
  RibbonCommandEvent,
  RibbonTabId,
  ViewTabHandles,
} from "./ribbon-types.js";
export { applyRibbonCommandToFlexSheet } from "./ribbon-flexsheet-hooks.js";
export { ViewRibbonController } from "./view-ribbon-controller.js";
