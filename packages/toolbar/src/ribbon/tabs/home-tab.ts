import {
  createToolbarButton,
  createToolbarDropdown,
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
  iconCopy,
  iconCut,
  iconDecreaseIndent,
  iconDoubleUnderline,
  iconFillColor,
  iconFontColor,
  iconFontGrow,
  iconFontShrink,
  iconFormatPainter,
  iconIncreaseIndent,
  iconItalic,
  iconMerge,
  iconPaste,
  iconRedo,
  iconTextOrientation,
  iconUnderline,
  iconUndo,
  iconWrapText,
} from "../../toolbar/icons.js";
import { mountAlignMergeMenu } from "../align-merge-menu.js";
import { mountAlignOrientationMenu } from "../align-orientation-menu.js";
import {
  RIBBON_FONT_FAMILY_DEFAULT_PREVIEW,
  RIBBON_FONT_FAMILY_ITEMS,
} from "../font-family-items.js";
import { mountFontBorderMenu } from "../font-border-menu.js";
import { mountRibbonColorPickerMenu } from "../ribbon-color-picker-menu.js";
import { createRibbonGroup } from "../ribbon-group.js";
import type { RibbonTabId } from "../ribbon-types.js";

const TAB: RibbonTabId = "home";

export interface HomeTabHandles {
  syncUndoRedo(canUndo: boolean, canRedo: boolean): void;
}

export function mountHomeTab(panel: HTMLElement, emit: RibbonEmit): HomeTabHandles {
  const inner = document.createElement("div");
  inner.className = "fs-ribbon-panel__inner";
  panel.appendChild(inner);

  // 撤销（剪贴板左侧：后退 / 撤销堆、前进 / 重做堆）
  const undoRedo = mountUndoGroup(inner, emit);

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
    rowTop.appendChild(
      createToolbarDropdown(
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
      ).element,
    );
    rowTop.appendChild(
      createToolbarDropdown(
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
      ).element,
    );
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
    rowActions.appendChild(createToolbarButton({ id: "home.font.bold", tab: TAB, label: "", icon: iconBold() }, emit).element);
    rowActions.appendChild(createToolbarButton({ id: "home.font.italic", tab: TAB, label: "", icon: iconItalic() }, emit).element);
    rowActions.appendChild(
      createToolbarButton({ id: "home.font.underline", tab: TAB, label: "", icon: iconUnderline() }, emit).element,
    );
    rowActions.appendChild(
      createToolbarButton(
        {
          id: "home.font.doubleUnderline",
          tab: TAB,
          label: "",
          icon: iconDoubleUnderline(),
          title: "双下划线",
        },
        emit,
      ).element,
    );
    const border = createToolbarButton(
      { id: "home.font.border", tab: TAB, label: "", icon: iconBorderAll(), splitDropdown: true, title: "边框" },
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
        icon: iconFillColor(),
        title: "填充颜色",
        colorPickerToggle: true,
      },
      emit,
    );
    fill.element.id = "fs-ribbon-home-font-fill";
    mountRibbonColorPickerMenu(fill.element, emit, TAB, "fill");
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
    mountRibbonColorPickerMenu(fontColor.element, emit, TAB, "font");
    rowActions.appendChild(fontColor.element);
    const stack = document.createElement("div");
    stack.className = "fs-ribbon-stack";
    stack.appendChild(rowTop);
    stack.appendChild(rowActions);
    content.appendChild(stack);
    inner.appendChild(root);
  }

  // 对齐方式
  {
    const { root, content } = createRibbonGroup("对齐方式");
    content.classList.add("fs-ribbon-align");

    const leftCol = document.createElement("div");
    leftCol.className = "fs-ribbon-align__left-col";

    const leftGrid = document.createElement("div");
    leftGrid.className = "fs-ribbon-align__grid";
    leftGrid.appendChild(createToolbarButton({ id: "home.align.top", tab: TAB, label: "", icon: iconAlignTop() }, emit).element);
    leftGrid.appendChild(
      createToolbarButton({ id: "home.align.middle", tab: TAB, label: "", icon: iconAlignMiddle() }, emit).element,
    );
    leftGrid.appendChild(
      createToolbarButton({ id: "home.align.bottom", tab: TAB, label: "", icon: iconAlignBottom() }, emit).element,
    );
    leftGrid.appendChild(createToolbarButton({ id: "home.align.left", tab: TAB, label: "", icon: iconAlignLeft() }, emit).element);
    leftGrid.appendChild(
      createToolbarButton({ id: "home.align.center", tab: TAB, label: "", icon: iconAlignCenter() }, emit).element,
    );
    leftGrid.appendChild(
      createToolbarButton({ id: "home.align.right", tab: TAB, label: "", icon: iconAlignRight() }, emit).element,
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
      createToolbarButton({ id: "home.align.wrap", tab: TAB, label: "自动换行", icon: iconWrapText() }, emit).element,
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

  return undoRedo;
}

function mountUndoGroup(inner: HTMLElement, emit: RibbonEmit): HomeTabHandles {
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
