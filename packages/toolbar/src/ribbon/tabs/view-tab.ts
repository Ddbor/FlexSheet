import {
  createToolbarButton,
  createToolbarDropdown,
  type RibbonEmit,
} from "../../toolbar/index.js";
import {
  iconFormulaBar,
  iconFreeze,
  iconGrid,
  iconNormalView,
  iconPageBreak,
  iconRecordMacro,
  iconRelativeRef,
  iconRuler,
  iconWindow,
  iconZoom,
} from "../../toolbar/icons.js";
import { createRibbonGroup } from "../ribbon-group.js";
import type { RibbonTabId } from "../ribbon-types.js";

const TAB: RibbonTabId = "view";

export interface ViewTabHandles {
  setZoomLabel(percent: number): void;
  setFreezeMenuLabel(text: string): void;
  setMacroRecordPressed(on: boolean): void;
  setMacroRelativePressed(on: boolean): void;
  setTogglePressed(id: string, pressed: boolean): void;
}

/**
 * 视图选项卡：缩放、冻结、显示项、窗口、宏；返回用于与 CanvasRenderer 状态同步的句柄。
 */
export function mountViewTab(panel: HTMLElement, emit: RibbonEmit): ViewTabHandles {
  const inner = document.createElement("div");
  inner.className = "fs-ribbon-panel__inner";
  panel.appendChild(inner);

  let zoomSetLabel: (text: string) => void = () => {};
  let freezeSetLabel: (text: string) => void = () => {};
  const toggleButtons = new Map<string, HTMLButtonElement>();

  {
    const { root, content } = createRibbonGroup("工作簿视图");
    content.appendChild(
      createToolbarButton(
        { id: "view.normal", tab: TAB, label: "普通", icon: iconNormalView(), variant: "large" },
        emit,
      ).element,
    );
    content.appendChild(
      createToolbarButton(
        { id: "view.pageBreakPreview", tab: TAB, label: "分页预览", icon: iconPageBreak(), variant: "large" },
        emit,
      ).element,
    );
    inner.appendChild(root);
  }

  {
    const { root, content } = createRibbonGroup("显示比例");
    const zoomDd = createToolbarDropdown(
      {
        id: "view.zoom",
        tab: TAB,
        label: "100%",
        title: "缩放比例",
        wide: true,
        items: [
          { id: "view.zoom.200", label: "200%" },
          { id: "view.zoom.100", label: "100%" },
          { id: "view.zoom.75", label: "75%" },
          { id: "view.zoom.50", label: "50%" },
          { id: "view.zoom.fitSelection", label: "缩放到所选区域" },
          { id: "view.zoom.custom", label: "自定义缩放…" },
        ],
      },
      emit,
    );
    zoomSetLabel = zoomDd.setLabel;
    content.appendChild(zoomDd.element);
    const row = document.createElement("div");
    row.className = "fs-ribbon-stack__row";
    row.appendChild(
      createToolbarButton({ id: "view.zoom.in", tab: TAB, label: "放大", icon: iconZoom() }, emit)
        .element,
    );
    row.appendChild(
      createToolbarButton({ id: "view.zoom.out", tab: TAB, label: "缩小", icon: iconZoom() }, emit)
        .element,
    );
    content.appendChild(row);
    inner.appendChild(root);
  }

  {
    const { root, content } = createRibbonGroup("显示");
    const row = document.createElement("div");
    row.className = "fs-ribbon-stack__row";
    const bRuler = createToolbarButton(
      {
        id: "view.toggle.ruler",
        tab: TAB,
        label: "标尺",
        icon: iconRuler(),
        title: "显示或隐藏标尺",
      },
      emit,
    );
    const bGrid = createToolbarButton(
      {
        id: "view.toggle.gridlines",
        tab: TAB,
        label: "网格线",
        icon: iconGrid(),
        title: "显示或隐藏网格线",
      },
      emit,
    );
    const bFormula = createToolbarButton(
      {
        id: "view.toggle.formulaBar",
        tab: TAB,
        label: "编辑栏",
        icon: iconFormulaBar(),
        title: "显示或隐藏编辑栏",
      },
      emit,
    );
    const bHead = createToolbarButton(
      {
        id: "view.toggle.headings",
        tab: TAB,
        label: "标题",
        icon: iconGrid(),
        title: "显示或隐藏行号与列标",
      },
      emit,
    );
    toggleButtons.set("view.toggle.ruler", bRuler.element);
    toggleButtons.set("view.toggle.gridlines", bGrid.element);
    toggleButtons.set("view.toggle.formulaBar", bFormula.element);
    toggleButtons.set("view.toggle.headings", bHead.element);
    row.appendChild(bRuler.element);
    row.appendChild(bGrid.element);
    row.appendChild(bFormula.element);
    row.appendChild(bHead.element);
    content.appendChild(row);
    inner.appendChild(root);
  }

  {
    const { root, content } = createRibbonGroup("冻结窗格");
    const freezeDd = createToolbarDropdown(
      {
        id: "view.freezePanes",
        tab: TAB,
        label: "冻结窗格",
        title: "冻结首行、首列或按活动单元格拆分",
        items: [
          { id: "view.freeze.none", label: "取消冻结" },
          { id: "view.freeze.split", label: "冻结拆分窗格" },
          { id: "view.freeze.topRow", label: "冻结首行" },
          { id: "view.freeze.firstCol", label: "冻结首列" },
        ],
      },
      emit,
    );
    freezeSetLabel = freezeDd.setLabel;
    const row = document.createElement("div");
    row.className = "fs-ribbon-stack__row";
    row.appendChild(freezeDd.element);
    row.appendChild(
      createToolbarButton(
        { id: "view.freeze.quick", tab: TAB, label: "冻结", icon: iconFreeze() },
        emit,
      ).element,
    );
    content.appendChild(row);
    inner.appendChild(root);
  }

  {
    const { root, content } = createRibbonGroup("窗口");
    const row = document.createElement("div");
    row.className = "fs-ribbon-stack__row";
    row.appendChild(
      createToolbarButton(
        { id: "view.window.new", tab: TAB, label: "新建窗口", icon: iconWindow() },
        emit,
      ).element,
    );
    row.appendChild(
      createToolbarButton(
        { id: "view.window.arrangeAll", tab: TAB, label: "全部重排", icon: iconWindow() },
        emit,
      ).element,
    );
    row.appendChild(
      createToolbarButton(
        { id: "view.window.switch", tab: TAB, label: "切换窗口", icon: iconWindow() },
        emit,
      ).element,
    );
    content.appendChild(row);
    inner.appendChild(root);
  }

  {
    const { root, content } = createRibbonGroup("宏");
    const row = document.createElement("div");
    row.className = "fs-ribbon-stack__row";
    const bRec = createToolbarButton(
      {
        id: "view.macro.record",
        tab: TAB,
        label: "录制宏",
        icon: iconRecordMacro(),
        title: "开始或停止录制宏",
      },
      emit,
    );
    const bRel = createToolbarButton(
      {
        id: "view.macro.relative",
        tab: TAB,
        label: "使用相对引用",
        icon: iconRelativeRef(),
        title: "宏中使用相对单元格引用",
      },
      emit,
    );
    bRec.element.setAttribute("aria-pressed", "false");
    bRel.element.setAttribute("aria-pressed", "false");
    toggleButtons.set("view.macro.record", bRec.element);
    toggleButtons.set("view.macro.relative", bRel.element);
    row.appendChild(bRec.element);
    row.appendChild(bRel.element);
    content.appendChild(row);
    inner.appendChild(root);
  }

  return {
    setZoomLabel(percent: number): void {
      zoomSetLabel(`${Math.round(percent)}%`);
    },
    setFreezeMenuLabel(text: string): void {
      freezeSetLabel(text);
    },
    setMacroRecordPressed(on: boolean): void {
      const el = toggleButtons.get("view.macro.record");
      if (el !== undefined) {
        el.setAttribute("aria-pressed", on ? "true" : "false");
      }
    },
    setMacroRelativePressed(on: boolean): void {
      const el = toggleButtons.get("view.macro.relative");
      if (el !== undefined) {
        el.setAttribute("aria-pressed", on ? "true" : "false");
      }
    },
    setTogglePressed(id: string, pressed: boolean): void {
      const el = toggleButtons.get(id);
      if (el !== undefined) {
        el.setAttribute("aria-pressed", pressed ? "true" : "false");
      }
    },
  };
}
