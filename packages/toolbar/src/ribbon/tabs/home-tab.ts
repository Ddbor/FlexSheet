import {
  createToolbarButton,
  createToolbarDropdown,
  type DropdownItem,
  type RibbonEmit,
} from "../../toolbar/index.js";
import {
  iconAlignBottom,
  iconAlignCenter,
  iconAlignLeft,
  iconAlignMiddle,
  iconAlignRight,
  iconAlignTop,
  iconBold,
  iconBorderAll,
  iconCellStyle,
  iconClear,
  iconCommaStyle,
  iconConditionalFormatting,
  iconCopy,
  iconCut,
  iconDecreaseDecimal,
  iconDecreaseIndent,
  iconDoubleUnderline,
  iconFillColor,
  iconFontFillColor,
  iconFontColor,
  iconFontGrow,
  iconFontShrink,
  iconFormatAccounting,
  iconFormatCurrency,
  iconFormatDate,
  iconFormatFraction,
  iconFormatGeneral,
  iconFormatNumber,
  iconFormatPainter,
  iconFormatScientific,
  iconFormatText,
  iconFormatTime,
  iconFunction,
  iconIncreaseDecimal,
  iconIncreaseIndent,
  iconItalic,
  iconMerge,
  iconPaste,
  iconPercent,
  iconRedo,
  iconTableStyle,
  iconTextOrientation,
  iconUnderline,
  iconUndo,
  iconWrapText,
} from "../../toolbar/icons.js";
import { mountAlignMergeMenu } from "../align-merge-menu.js";
import { mountAlignOrientationMenu } from "../align-orientation-menu.js";
import { mountCellStyleRibbonMenu } from "../cell-styles-ribbon-menu.js";
import { mountHomeClearMenu } from "../home-clear-menu.js";
import { mountAutoSumSubmenu } from "../home-autosum-menu.js";
import { mountHomeFindMenu } from "../home-find-menu.js";
import { mountHomeSortFilterMenu } from "../home-sort-filter-menu.js";
import { mountHomeFillMenu } from "../home-fill-menu.js";
import { mountConditionalFormatMenu, mountTableFormatStyleMenu } from "../home-styles-menus.js";
import {
  RIBBON_FONT_FAMILY_DEFAULT_PREVIEW,
  RIBBON_FONT_FAMILY_ITEMS,
} from "../font-family-items.js";
import { mountFontBorderMenu } from "../font-border-menu.js";
import { mountRibbonColorPickerMenu } from "../ribbon-color-picker-menu.js";
import type { RibbonHomeFontChromeState } from "../ribbon-font-chrome.js";
import {
  RIBBON_NUMBER_FORMAT_PRESETS,
  type RibbonHomeNumberFormatChromeState,
} from "../ribbon-number-format-chrome.js";

/** 与 `RIBBON_NUMBER_FORMAT_PRESETS` 各项一一对应（icons.ts 数字格式图标） */
const NUMBER_FORMAT_DROPDOWN_ICONS: Record<
  (typeof RIBBON_NUMBER_FORMAT_PRESETS)[number]["id"],
  () => SVGSVGElement
> = {
  "home.number.format.general": () => iconFormatGeneral(),
  "home.number.format.number": () => iconFormatNumber(),
  "home.number.format.currency": () => iconFormatCurrency(),
  "home.number.format.accounting": () => iconFormatAccounting(),
  "home.number.format.shortDate": () => iconFormatDate(),
  "home.number.format.longDate": () => iconFormatDate(),
  "home.number.format.time": () => iconFormatTime(),
  "home.number.format.percent": () => iconPercent(),
  "home.number.format.fraction": () => iconFormatFraction(),
  "home.number.format.scientific": () => iconFormatScientific(),
  "home.number.format.text": () => iconFormatText(),
};
import { createRibbonGroup } from "../ribbon-group.js";
import type { FlexSheetLike, RibbonTabId } from "../ribbon-types.js";
import { iconFindRibbon } from "../home-find-icons.js";
import { iconSortFilterRibbon } from "../home-sort-filter-icons.js";

const TAB: RibbonTabId = "home";

export interface HomeTabHandles {
  syncUndoRedo(canUndo: boolean, canRedo: boolean): void;
  /** 与活动单元格字体样式同步（字体名、字号、加粗/斜体/下划线状态）。 */
  syncFontChrome(state: RibbonHomeFontChromeState): void;
  /** 与活动单元格数字格式分类下拉同步。 */
  syncNumberFormatChrome(state: RibbonHomeNumberFormatChromeState): void;
}

export function mountHomeTab(
  panel: HTMLElement,
  emit: RibbonEmit,
  getFlexSheet?: () => FlexSheetLike | undefined,
): HomeTabHandles {
  const inner = document.createElement("div");
  inner.className = "fs-ribbon-panel__inner";
  panel.appendChild(inner);

  // 撤销（剪贴板左侧：后退 / 撤销堆、前进 / 重做堆）
  const undoRedo = mountUndoGroup(inner, emit);
  let syncFontChrome: HomeTabHandles["syncFontChrome"] = (): void => {};
  let syncNumberFormatChrome: HomeTabHandles["syncNumberFormatChrome"] = (): void => {};

  // 剪贴板
  {
    const { root, content } = createRibbonGroup("剪贴板");
    content.classList.add("fs-ribbon-clipboard");
    const paste = createToolbarButton(
      {
        id: "home.clipboard.paste",
        tab: TAB,
        label: "粘贴",
        variant: "large",
        icon: iconPaste(),
      },
      emit,
    );
    const side = document.createElement("div");
    side.className = "fs-ribbon-clipboard__side";
    const cut = createToolbarButton(
      { id: "home.clipboard.cut", tab: TAB, label: "", icon: iconCut(), title: "剪切" },
      emit,
    );
    cut.element.classList.add("fs-ribbon-clipboard__item");
    side.appendChild(cut.element);
    const copy = createToolbarButton(
      { id: "home.clipboard.copy", tab: TAB, label: "", icon: iconCopy(), title: "复制" },
      emit,
    );
    copy.element.classList.add("fs-ribbon-clipboard__item");
    side.appendChild(copy.element);
    const format = createToolbarButton(
      {
        id: "home.clipboard.formatPainter",
        tab: TAB,
        label: "",
        icon: iconFormatPainter(),
        title: "格式刷",
      },
      emit,
    );
    format.element.classList.add("fs-ribbon-clipboard__item");
    side.appendChild(format.element);
    content.appendChild(paste.element);
    content.appendChild(side);
    inner.appendChild(root);
  }

  // 字体
  {
    const { root, content } = createRibbonGroup("字体");
    content.classList.add("fs-ribbon-font");
    const rowTop = document.createElement("div");
    rowTop.className = "fs-ribbon-stack__row fs-ribbon-stack__row--gap";
    const fontFamilyDd = createToolbarDropdown(
      {
        id: "home.font.family",
        tab: TAB,
        label: "微软雅黑",
        wide: true,
        menuClassName: "fs-dd__menu--font-list",
        initialLabelFontFamily: RIBBON_FONT_FAMILY_DEFAULT_PREVIEW,
        items: RIBBON_FONT_FAMILY_ITEMS,
      },
      emit,
    );
    rowTop.appendChild(fontFamilyDd.element);
    const fontSizeDd = createToolbarDropdown(
      {
        id: "home.font.size",
        tab: TAB,
        label: "11",
        items: [
          { id: "home.font.size.8", label: "8" },
          { id: "home.font.size.9", label: "9" },
          { id: "home.font.size.10", label: "10" },
          { id: "home.font.size.11", label: "11" },
          { id: "home.font.size.12", label: "12" },
          { id: "home.font.size.14", label: "14" },
          { id: "home.font.size.16", label: "16" },
          { id: "home.font.size.18", label: "18" },
          { id: "home.font.size.20", label: "20" },
          { id: "home.font.size.24", label: "24" },
          { id: "home.font.size.26", label: "26" },
          { id: "home.font.size.28", label: "28" },
          { id: "home.font.size.36", label: "36" },
          { id: "home.font.size.48", label: "48" },
          { id: "home.font.size.72", label: "72" },
        ],
      },
      emit,
    );
    rowTop.appendChild(fontSizeDd.element);
    rowTop.appendChild(
      createToolbarButton(
        { id: "home.font.grow", tab: TAB, label: "", icon: iconFontGrow(), title: "增大字号" },
        emit,
      ).element,
    );
    rowTop.appendChild(
      createToolbarButton(
        { id: "home.font.shrink", tab: TAB, label: "", icon: iconFontShrink(), title: "减小字号" },
        emit,
      ).element,
    );
    const rowActions = document.createElement("div");
    rowActions.className = "fs-ribbon-font__grid";
    const boldBtn = createToolbarButton(
      { id: "home.font.bold", tab: TAB, label: "", icon: iconBold() },
      emit,
    );
    const italicBtn = createToolbarButton(
      { id: "home.font.italic", tab: TAB, label: "", icon: iconItalic() },
      emit,
    );
    const underlineBtn = createToolbarButton(
      { id: "home.font.underline", tab: TAB, label: "", icon: iconUnderline() },
      emit,
    );
    const doubleUnderlineBtn = createToolbarButton(
      {
        id: "home.font.doubleUnderline",
        tab: TAB,
        label: "",
        icon: iconDoubleUnderline(),
        title: "双下划线",
      },
      emit,
    );
    rowActions.appendChild(boldBtn.element);
    rowActions.appendChild(italicBtn.element);
    rowActions.appendChild(underlineBtn.element);
    rowActions.appendChild(doubleUnderlineBtn.element);
    const border = createToolbarButton(
      {
        id: "home.font.border",
        tab: TAB,
        label: "",
        icon: iconBorderAll(),
        splitDropdown: true,
        title: "边框",
      },
      emit,
    );
    border.element.id = "fs-ribbon-home-font-border";
    mountFontBorderMenu(border.element, emit, TAB);
    rowActions.appendChild(border.element);
    const fill = createToolbarButton(
      {
        id: "home.font.fill",
        tab: TAB,
        label: "",
        icon: iconFontFillColor(),
        title: "填充颜色",
        colorPickerToggle: true,
      },
      emit,
    );
    fill.element.id = "fs-ribbon-home-font-fill";
    fill.element.classList.add("fs-tb-btn--colorbar");
    mountRibbonColorPickerMenu(fill.element, emit, TAB, "fill");
    const fillColorBar = document.createElement("span");
    fillColorBar.className = "fs-tb-btn__colorbar";
    fill.element.appendChild(fillColorBar);
    rowActions.appendChild(fill.element);
    const fontColor = createToolbarButton(
      {
        id: "home.font.color",
        tab: TAB,
        label: "",
        icon: iconFontColor(),
        title: "字体颜色",
        colorPickerToggle: true,
      },
      emit,
    );
    fontColor.element.id = "fs-ribbon-home-font-color";
    fontColor.element.classList.add("fs-tb-btn--colorbar");
    mountRibbonColorPickerMenu(fontColor.element, emit, TAB, "font");
    const fontColorBar = document.createElement("span");
    fontColorBar.className = "fs-tb-btn__colorbar";
    fontColor.element.appendChild(fontColorBar);
    rowActions.appendChild(fontColor.element);
    const stack = document.createElement("div");
    stack.className = "fs-ribbon-stack";
    stack.appendChild(rowTop);
    stack.appendChild(rowActions);
    content.appendChild(stack);
    inner.appendChild(root);

    syncFontChrome = (state: RibbonHomeFontChromeState): void => {
      fontFamilyDd.setLabel(state.fontLabel, state.fontPreviewCss);
      fontSizeDd.setLabel(state.sizeLabel);
      boldBtn.setPressed(state.boldPressed);
      italicBtn.setPressed(state.italicPressed);
      underlineBtn.setPressed(state.underlinePressed);
      doubleUnderlineBtn.setPressed(state.doubleUnderlinePressed);
      const emptyFill = state.fillStripeCss === null;
      fillColorBar.classList.toggle("fs-tb-btn__colorbar--empty", emptyFill);
      fillColorBar.style.backgroundColor = emptyFill ? "" : state.fillStripeCss;
      fontColorBar.classList.remove("fs-tb-btn__colorbar--empty");
      fontColorBar.style.backgroundColor = state.fontStripeCss;
    };
  }

  // 数字
  {
    const { root, content } = createRibbonGroup("数字");
    content.classList.add("fs-ribbon-number");
    const stack = document.createElement("div");
    stack.className = "fs-ribbon-stack";
    const rowDd = document.createElement("div");
    rowDd.className = "fs-ribbon-stack__row fs-ribbon-stack__row--gap";
    const presetItems: readonly DropdownItem[] = RIBBON_NUMBER_FORMAT_PRESETS.map((p) => ({
      id: p.id,
      label: p.label,
      icon: NUMBER_FORMAT_DROPDOWN_ICONS[p.id],
    }));
    const numberFormatDd = createToolbarDropdown(
      {
        id: "home.number.format.dropdown",
        tab: TAB,
        label: "常规",
        wide: true,
        title: "数字格式",
        items: presetItems,
      },
      emit,
    );
    rowDd.appendChild(numberFormatDd.element);
    stack.appendChild(rowDd);
    const quick = document.createElement("div");
    quick.className = "fs-ribbon-number__quick";
    quick.appendChild(
      createToolbarButton(
        {
          id: "home.number.quick.percent",
          tab: TAB,
          label: "",
          icon: iconPercent(),
          title: "百分比样式",
        },
        emit,
      ).element,
    );
    quick.appendChild(
      createToolbarButton(
        {
          id: "home.number.quick.comma",
          tab: TAB,
          label: "",
          icon: iconCommaStyle(),
          title: "千位分隔样式",
        },
        emit,
      ).element,
    );
    const sep = document.createElement("div");
    sep.className = "fs-ribbon-number__quick-sep";
    sep.setAttribute("aria-hidden", "true");
    quick.appendChild(sep);
    quick.appendChild(
      createToolbarButton(
        {
          id: "home.number.quick.increaseDecimal",
          tab: TAB,
          label: "",
          icon: iconIncreaseDecimal(),
          title: "增加小数位数",
        },
        emit,
      ).element,
    );
    quick.appendChild(
      createToolbarButton(
        {
          id: "home.number.quick.decreaseDecimal",
          tab: TAB,
          label: "",
          icon: iconDecreaseDecimal(),
          title: "减少小数位数",
        },
        emit,
      ).element,
    );
    stack.appendChild(quick);
    content.appendChild(stack);
    inner.appendChild(root);

    syncNumberFormatChrome = (state: RibbonHomeNumberFormatChromeState): void => {
      numberFormatDd.setLabel(state.categoryLabel);
    };
  }

  // 对齐方式
  {
    const { root, content } = createRibbonGroup("对齐方式");
    content.classList.add("fs-ribbon-align");

    const leftCol = document.createElement("div");
    leftCol.className = "fs-ribbon-align__left-col";

    const leftGrid = document.createElement("div");
    leftGrid.className = "fs-ribbon-align__grid";
    leftGrid.appendChild(
      createToolbarButton(
        { id: "home.align.top", tab: TAB, label: "", icon: iconAlignTop(), title: "顶端对齐" },
        emit,
      ).element,
    );
    leftGrid.appendChild(
      createToolbarButton(
        {
          id: "home.align.middle",
          tab: TAB,
          label: "",
          icon: iconAlignMiddle(),
          title: "垂直居中",
        },
        emit,
      ).element,
    );
    leftGrid.appendChild(
      createToolbarButton(
        {
          id: "home.align.bottom",
          tab: TAB,
          label: "",
          icon: iconAlignBottom(),
          title: "底端对齐",
        },
        emit,
      ).element,
    );
    leftGrid.appendChild(
      createToolbarButton(
        { id: "home.align.left", tab: TAB, label: "", icon: iconAlignLeft(), title: "左对齐" },
        emit,
      ).element,
    );
    leftGrid.appendChild(
      createToolbarButton(
        { id: "home.align.center", tab: TAB, label: "", icon: iconAlignCenter(), title: "居中" },
        emit,
      ).element,
    );
    leftGrid.appendChild(
      createToolbarButton(
        { id: "home.align.right", tab: TAB, label: "", icon: iconAlignRight(), title: "右对齐" },
        emit,
      ).element,
    );
    leftCol.appendChild(leftGrid);

    const indentRow = document.createElement("div");
    indentRow.className = "fs-ribbon-align__indent-row";
    indentRow.appendChild(
      createToolbarButton(
        {
          id: "home.align.indentDecrease",
          tab: TAB,
          label: "",
          icon: iconDecreaseIndent(),
          title: "减少缩进量",
        },
        emit,
      ).element,
    );
    indentRow.appendChild(
      createToolbarButton(
        {
          id: "home.align.indentIncrease",
          tab: TAB,
          label: "",
          icon: iconIncreaseIndent(),
          title: "增加缩进量",
        },
        emit,
      ).element,
    );
    leftCol.appendChild(indentRow);

    const rightStack = document.createElement("div");
    rightStack.className = "fs-ribbon-align__right";
    rightStack.appendChild(
      createToolbarButton(
        {
          id: "home.align.wrap",
          tab: TAB,
          label: "自动换行",
          icon: iconWrapText(),
          title: "自动换行",
        },
        emit,
      ).element,
    );
    const merge = createToolbarButton(
      {
        id: "home.align.merge",
        tab: TAB,
        label: "合并后居中",
        icon: iconMerge(),
        splitDropdown: true,
        title: "合并后居中",
      },
      emit,
    );
    merge.element.id = "fs-ribbon-home-align-merge";
    mountAlignMergeMenu(merge.element, emit, TAB);
    rightStack.appendChild(merge.element);

    const orientation = createToolbarButton(
      {
        id: "home.align.textDirection.counterClockwise",
        tab: TAB,
        label: "方向",
        icon: iconTextOrientation(),
        splitDropdown: true,
        title: "方向",
      },
      emit,
    );
    orientation.element.id = "fs-ribbon-home-align-orientation";
    mountAlignOrientationMenu(orientation.element, emit, TAB);
    rightStack.appendChild(orientation.element);

    content.appendChild(leftCol);
    content.appendChild(rightStack);
    inner.appendChild(root);
  }

  // 样式（条件格式 / 套用表格格式 / 单元格样式）
  {
    const { root, content } = createRibbonGroup("样式");
    content.classList.add("fs-ribbon-styles");
    const conditionalBtn = createToolbarButton(
      {
        id: "home.style.conditional",
        tab: TAB,
        label: "条件格式",
        variant: "large",
        icon: iconConditionalFormatting(),
        menuTrigger: true,
        title: "条件格式",
      },
      emit,
    );
    conditionalBtn.element.id = "fs-ribbon-home-style-conditional";
    mountConditionalFormatMenu(conditionalBtn.element, emit, TAB);
    const tableFormatBtn = createToolbarButton(
      {
        id: "home.style.tableFormat",
        tab: TAB,
        label: "套用表格格式",
        variant: "large",
        icon: iconTableStyle(),
        menuTrigger: true,
        title: "套用表格格式",
      },
      emit,
    );
    tableFormatBtn.element.id = "fs-ribbon-home-style-table-format";
    mountTableFormatStyleMenu(tableFormatBtn.element, emit, TAB, {
      getCustomTableStyles: () => getFlexSheet?.()?.getCustomTableStyleEntries?.() ?? [],
    });
    const cellStylesBtn = createToolbarButton(
      {
        id: "home.style.cellStyles",
        tab: TAB,
        label: "单元格样式",
        variant: "large",
        icon: iconCellStyle(),
        menuTrigger: true,
        title: "单元格样式",
      },
      emit,
    );
    cellStylesBtn.element.id = "fs-ribbon-home-style-cell-styles";
    mountCellStyleRibbonMenu(cellStylesBtn.element, emit, TAB);
    const sep = document.createElement("div");
    sep.className = "fs-ribbon-styles__sep";
    sep.setAttribute("aria-hidden", "true");
    const cellsOps = document.createElement("div");
    cellsOps.className = "fs-ribbon-styles__cells-ops";
    const makeCellOpBtn = (id: string, label: string, icon: SVGSVGElement): HTMLButtonElement => {
      const btn = createToolbarButton(
        {
          id,
          tab: TAB,
          label,
          icon,
          menuTrigger: true,
          title: label,
        },
        emit,
      );
      return btn.element;
    };
    {
      const sumBtn = createToolbarButton(
        {
          id: "home.cells.autoSum",
          tab: TAB,
          label: "自动求和",
          icon: iconFunction(),
          splitDropdown: true,
          title: "自动求和 / 其他函数",
        },
        emit,
      );
      sumBtn.element.id = "fs-ribbon-home-cells-autosum";
      mountAutoSumSubmenu(sumBtn.element, emit, TAB);
      cellsOps.appendChild(sumBtn.element);
    }
    const fillBtn = makeCellOpBtn("home.cells.fill", "填充", iconFillColor());
    fillBtn.id = "fs-ribbon-home-cells-fill";
    mountHomeFillMenu(fillBtn, emit, TAB);
    cellsOps.appendChild(fillBtn);
    const clearBtn = makeCellOpBtn("home.cells.clear", "清除", iconClear());
    clearBtn.id = "fs-ribbon-home-cells-clear";
    mountHomeClearMenu(clearBtn, emit, TAB);
    cellsOps.appendChild(clearBtn);
    content.appendChild(conditionalBtn.element);
    content.appendChild(tableFormatBtn.element);
    content.appendChild(cellStylesBtn.element);
    content.appendChild(sep);
    content.appendChild(cellsOps);
    inner.appendChild(root);
  }

  // 排序和筛选（面板最右）
  {
    const { root, content } = createRibbonGroup("排序和筛选");
    content.classList.add("fs-ribbon-sort-filter");
    const sortFilterBtn = createToolbarButton(
      {
        id: "home.sortFilter",
        tab: TAB,
        label: "排序和筛选",
        variant: "large",
        icon: iconSortFilterRibbon(),
        menuTrigger: true,
        title: "排序和筛选",
      },
      emit,
    );
    sortFilterBtn.element.id = "fs-ribbon-home-sort-filter";
    mountHomeSortFilterMenu(sortFilterBtn.element, emit, TAB);
    content.appendChild(sortFilterBtn.element);
    inner.appendChild(root);
  }

  {
    const { root, content } = createRibbonGroup("查找");
    content.classList.add("fs-ribbon-find");
    const findBtn = createToolbarButton(
      {
        id: "home.find",
        tab: TAB,
        label: "查找",
        variant: "large",
        icon: iconFindRibbon(),
        menuTrigger: true,
        title: "查找",
      },
      emit,
    );
    findBtn.element.id = "fs-ribbon-home-find";
    mountHomeFindMenu(findBtn.element, emit, TAB);
    content.appendChild(findBtn.element);
    inner.appendChild(root);
  }

  return {
    syncUndoRedo: undoRedo.syncUndoRedo,
    syncFontChrome,
    syncNumberFormatChrome,
  };
}

function mountUndoGroup(
  inner: HTMLElement,
  emit: RibbonEmit,
): Pick<HomeTabHandles, "syncUndoRedo"> {
  const { root, content } = createRibbonGroup("撤销");
  content.classList.add("fs-ribbon-undo");
  const back = createToolbarButton(
    {
      id: "home.undo.back",
      tab: TAB,
      label: "",
      icon: iconUndo(),
      title: "后退 (Ctrl+Z)",
      disabled: true,
    },
    emit,
  );
  const forward = createToolbarButton(
    {
      id: "home.undo.forward",
      tab: TAB,
      label: "",
      icon: iconRedo(),
      title: "前进 (Ctrl+Y)",
      disabled: true,
    },
    emit,
  );
  const stack = document.createElement("div");
  stack.className = "fs-ribbon-undo__stack";
  stack.appendChild(back.element);
  stack.appendChild(forward.element);
  content.appendChild(stack);
  inner.appendChild(root);

  return {
    syncUndoRedo(canUndo: boolean, canRedo: boolean): void {
      back.setDisabled(!canUndo);
      forward.setDisabled(!canRedo);
    },
  };
}
