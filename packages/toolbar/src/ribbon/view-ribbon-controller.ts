import type { SelectionRange } from "@flexsheet/core";
import type { FlexSheetLike, RibbonCommandEvent } from "./ribbon-types.js";
import type { ViewTabHandles } from "./tabs/view-tab.js";

const CUSTOM_ZOOM_MIN = 10;
const CUSTOM_ZOOM_MAX = 400;

/**
 * 将 Ribbon「视图」选项卡命令与 `CanvasRenderer` / `FlexSheet` 联动，并维护 Ribbon 控件状态（缩放标签、冻结菜单、开关与宏）。
 */
export class ViewRibbonController {
  private readonly flexSheet: FlexSheetLike;
  private readonly viewTab: ViewTabHandles;
  private readonly onMacroRecord?: (recording: boolean) => void;

  constructor(options: {
    readonly flexSheet: FlexSheetLike;
    readonly viewTabHandles: ViewTabHandles;
    readonly onMacroRecord?: (recording: boolean) => void;
  }) {
    this.flexSheet = options.flexSheet;
    this.viewTab = options.viewTabHandles;
    this.onMacroRecord = options.onMacroRecord;
  }

  /** 从当前渲染器同步 Ribbon 标签与按钮状态（主题切换或外部改 renderer 后调用）。 */
  syncFromRenderer(): void {
    const r = this.flexSheet.getRenderer();
    this.viewTab.setZoomLabel(r.getViewZoom() * 100);
    const fr = r.frozenRows;
    const fc = r.frozenCols;
    if (fr === 0 && fc === 0) {
      this.viewTab.setFreezeMenuLabel("冻结窗格");
    } else {
      this.viewTab.setFreezeMenuLabel(`已冻结 (${fr}×${fc})`);
    }
    this.viewTab.setMacroRecordPressed(r.macroRecording);
    this.viewTab.setMacroRelativePressed(r.macroUseRelativeReference);
    this.viewTab.setTogglePressed("view.toggle.gridlines", r.showGridLines);
    this.viewTab.setTogglePressed("view.toggle.headings", r.showHeadings);
    this.viewTab.setTogglePressed("view.toggle.ruler", r.showRuler);
    const fb = this.flexSheet.isFormulaBarVisible?.() ?? true;
    this.viewTab.setTogglePressed("view.toggle.formulaBar", fb);
  }

  /**
   * 处理视图相关命令；若已处理返回 true。
   */
  handleCommand(ev: RibbonCommandEvent): boolean {
    if (ev.tab !== "view") {
      return false;
    }
    const fs = this.flexSheet;
    const r = fs.getRenderer();

    switch (ev.id) {
      case "view.zoom.100":
        r.resetZoom100();
        break;
      case "view.zoom.200":
        r.setViewZoom(2);
        break;
      case "view.zoom.75":
        r.setViewZoom(0.75);
        break;
      case "view.zoom.50":
        r.setViewZoom(0.5);
        break;
      case "view.zoom.fitSelection": {
        const range = fs.selection.getNormalizedRange();
        r.zoomToFitRange(range);
        break;
      }
      case "view.zoom.custom": {
        const raw = window.prompt(
          "请输入缩放比例（10%–400%）",
          String(Math.round(r.getViewZoom() * 100)),
        );
        if (raw === null) {
          return true;
        }
        const n = Number.parseFloat(raw.replace(/%/g, "").trim());
        if (!Number.isFinite(n)) {
          return true;
        }
        const pct = n > 0 && n <= 4 ? n * 100 : n;
        const clamped = Math.max(CUSTOM_ZOOM_MIN, Math.min(CUSTOM_ZOOM_MAX, pct));
        r.setViewZoom(clamped / 100);
        break;
      }
      case "view.zoom.in":
        r.zoomIn();
        break;
      case "view.zoom.out":
        r.zoomOut();
        break;
      case "view.toggle.gridlines":
        r.setShowGridLines(!r.showGridLines);
        break;
      case "view.toggle.headings":
        r.setShowHeadings(!r.showHeadings);
        break;
      case "view.toggle.ruler":
        r.setShowRuler(!r.showRuler);
        break;
      case "view.toggle.formulaBar": {
        const cur = fs.isFormulaBarVisible?.() ?? true;
        fs.setFormulaBarVisible?.(!cur);
        break;
      }
      case "view.freeze.none":
        fs.setFrozenPanes(0, 0);
        break;
      case "view.freeze.topRow":
        fs.setFrozenPanes(1, 0);
        break;
      case "view.freeze.firstCol":
        fs.setFrozenPanes(0, 1);
        break;
      case "view.freeze.split": {
        const cell = fs.selection.getActiveCell();
        fs.setFrozenPanes(cell.row, cell.col);
        break;
      }
      case "view.freeze.quick": {
        const cell = fs.selection.getActiveCell();
        if (cell.row === 0 && cell.col === 0) {
          fs.setFrozenPanes(1, 1);
        } else {
          fs.setFrozenPanes(cell.row, cell.col);
        }
        break;
      }
      case "view.window.new": {
        window.open(window.location.href, "_blank", "noopener,noreferrer");
        break;
      }
      case "view.window.arrangeAll": {
        try {
          const sw = window.screen.availWidth;
          const sh = window.screen.availHeight;
          window.resizeTo(Math.floor(sw / 2), sh);
          window.moveTo(0, 0);
        } catch {
          /* ignore */
        }
        break;
      }
      case "view.window.switch": {
        if (window.opener !== null && !window.opener.closed) {
          try {
            window.opener.focus();
          } catch {
            /* ignore */
          }
        } else {
          window.alert("当前没有可切换的其他窗口（请从「新建窗口」打开多实例）。");
        }
        break;
      }
      case "view.macro.record": {
        const next = !r.macroRecording;
        r.setMacroRecording(next);
        this.onMacroRecord?.(next);
        break;
      }
      case "view.macro.relative": {
        r.setMacroUseRelativeReference(!r.macroUseRelativeReference);
        break;
      }
      case "view.normal":
      case "view.pageBreakPreview":
        window.console.info("[Ribbon] 视图模式占位:", ev.id);
        break;
      default:
        return false;
    }

    fs.refresh();
    this.syncFromRenderer();
    return true;
  }

  /** 供外部在选区或缩放变化后刷新标签（不重绘）。 */
  updateZoomLabelOnly(): void {
    const r = this.flexSheet.getRenderer();
    this.viewTab.setZoomLabel(r.getViewZoom() * 100);
  }

  /** 缩放到指定选区（与 `view.zoom.fitSelection` 相同逻辑）。 */
  zoomToRange(range: SelectionRange): void {
    const r = this.flexSheet.getRenderer();
    r.zoomToFitRange(range);
    this.flexSheet.refresh();
    this.syncFromRenderer();
  }
}
