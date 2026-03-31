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
  iconCommaStyle,
  iconCopy,
  iconCut,
  iconFillColor,
  iconFormatPainter,
  iconFontColor,
  iconItalic,
  iconMerge,
  iconOrientation,
  iconPaste,
  iconPercent,
  iconRedo,
  iconTable,
  iconUnderline,
  iconUndo,
  iconWrapText,
} from "../../toolbar/icons.js";
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
        splitDropdown: true,
      },
      emit,
    );
    paste.element.addEventListener("fs-dropdown-toggle", () => {
      emit("home.clipboard.pasteMenu", TAB);
    });
    const side = document.createElement("div");
    side.className = "fs-ribbon-clipboard__side";
    const cut = createToolbarButton({ id: "home.clipboard.cut", tab: TAB, label: "剪切", icon: iconCut() }, emit);
    cut.element.classList.add("fs-ribbon-clipboard__item");
    side.appendChild(cut.element);
    const copy = createToolbarButton({ id: "home.clipboard.copy", tab: TAB, label: "复制", icon: iconCopy() }, emit);
    copy.element.classList.add("fs-ribbon-clipboard__item");
    side.appendChild(copy.element);
    const format = createToolbarButton(
      { id: "home.clipboard.formatPainter", tab: TAB, label: "格式", icon: iconFormatPainter() },
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
          items: [
            { id: "home.font.family.segoe", label: "Segoe UI" },
            { id: "home.font.family.msYahei", label: "微软雅黑" },
            { id: "home.font.family.msSong", label: "宋体" },
            { id: "home.font.family.consolas", label: "Consolas" },
          ],
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
            { id: "home.font.size.18", label: "18" },
          ],
        },
        emit,
      ).element,
    );
    rowTop.appendChild(
      createToolbarButton({ id: "home.font.grow", tab: TAB, label: "A^", title: "增大字号" }, emit).element,
    );
    rowTop.appendChild(
      createToolbarButton({ id: "home.font.shrink", tab: TAB, label: "A˅", title: "减小字号" }, emit).element,
    );
    const rowActions = document.createElement("div");
    rowActions.className = "fs-ribbon-font__grid";
    rowActions.appendChild(createToolbarButton({ id: "home.font.bold", tab: TAB, label: "", icon: iconBold() }, emit).element);
    rowActions.appendChild(createToolbarButton({ id: "home.font.italic", tab: TAB, label: "", icon: iconItalic() }, emit).element);
    rowActions.appendChild(
      createToolbarButton({ id: "home.font.underline", tab: TAB, label: "", icon: iconUnderline() }, emit).element,
    );
    const border = createToolbarButton(
      { id: "home.font.border", tab: TAB, label: "", icon: iconTable(), splitDropdown: true },
      emit,
    );
    border.element.addEventListener("fs-dropdown-toggle", () => {
      emit("home.font.borderMenu", TAB);
    });
    rowActions.appendChild(border.element);
    const fill = createToolbarButton(
      { id: "home.font.fill", tab: TAB, label: "", icon: iconFillColor() },
      emit,
    );
    rowActions.appendChild(fill.element);
    const fontColor = createToolbarButton(
      { id: "home.font.color", tab: TAB, label: "", icon: iconFontColor() },
      emit,
    );
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

    const midStack = document.createElement("div");
    midStack.className = "fs-ribbon-align__mid";
    const orientation = createToolbarButton(
      { id: "home.align.orientation", tab: TAB, label: "", icon: iconOrientation(), splitDropdown: true },
      emit,
    );
    orientation.element.addEventListener("fs-dropdown-toggle", () => {
      emit("home.align.orientationMenu", TAB);
    });
    midStack.appendChild(orientation.element);
    const horizontal = createToolbarButton(
      { id: "home.align.horizontal", tab: TAB, label: "", icon: iconAlignLeft(), splitDropdown: true },
      emit,
    );
    horizontal.element.addEventListener("fs-dropdown-toggle", () => {
      emit("home.align.horizontalMenu", TAB);
    });
    midStack.appendChild(horizontal.element);

    const rightStack = document.createElement("div");
    rightStack.className = "fs-ribbon-align__right";
    rightStack.appendChild(
      createToolbarButton({ id: "home.align.wrap", tab: TAB, label: "自动换行", icon: iconWrapText() }, emit).element,
    );
    rightStack.appendChild(
      createToolbarButton({ id: "home.align.merge", tab: TAB, label: "合并后居中", icon: iconMerge() }, emit).element,
    );

    content.appendChild(leftGrid);
    content.appendChild(midStack);
    content.appendChild(rightStack);
    inner.appendChild(root);
  }

  // 数字
  {
    const { root, content } = createRibbonGroup("数字");
    const { element: fmt } = createToolbarDropdown(
      {
        id: "home.number.format",
        tab: TAB,
        label: "常规",
        wide: true,
        items: [
          { id: "home.number.general", label: "常规" },
          { id: "home.number.number", label: "数字" },
          { id: "home.number.currency", label: "货币" },
          { id: "home.number.percent", label: "百分比" },
          { id: "home.number.date", label: "日期" },
        ],
      },
      emit,
    );
    const row = document.createElement("div");
    row.className = "fs-ribbon-stack__row";
    row.appendChild(createToolbarButton({ id: "home.number.percentBtn", tab: TAB, label: "%", icon: iconPercent() }, emit).element);
    row.appendChild(createToolbarButton({ id: "home.number.comma", tab: TAB, label: "千分位", icon: iconCommaStyle() }, emit).element);
    const stack = document.createElement("div");
    stack.className = "fs-ribbon-stack";
    stack.appendChild(fmt);
    stack.appendChild(row);
    content.appendChild(stack);
    inner.appendChild(root);
  }

  // 样式
  {
    const { root, content } = createRibbonGroup("样式");
    const row = document.createElement("div");
    row.className = "fs-ribbon-stack__row";
    row.appendChild(
      createToolbarButton({ id: "home.style.conditional", tab: TAB, label: "条件格式", variant: "large" }, emit)
        .element,
    );
    row.appendChild(
      createToolbarButton({ id: "home.style.asTable", tab: TAB, label: "表格格式", variant: "large" }, emit)
        .element,
    );
    row.appendChild(
      createToolbarButton({ id: "home.style.cellStyles", tab: TAB, label: "单元格样式", variant: "large" }, emit)
        .element,
    );
    content.appendChild(row);
    inner.appendChild(root);
  }

  // 单元格
  {
    const { root, content } = createRibbonGroup("单元格");
    const stack = document.createElement("div");
    stack.className = "fs-ribbon-stack";
    const row1 = document.createElement("div");
    row1.className = "fs-ribbon-stack__row";
    row1.appendChild(
      createToolbarDropdown(
        {
          id: "home.cells.insert",
          tab: TAB,
          label: "插入",
          items: [
            { id: "home.cells.insert.cells", label: "插入单元格" },
            { id: "home.cells.insert.rows", label: "插入工作表行" },
            { id: "home.cells.insert.cols", label: "插入工作表列" },
          ],
        },
        emit,
      ).element,
    );
    row1.appendChild(
      createToolbarDropdown(
        {
          id: "home.cells.delete",
          tab: TAB,
          label: "删除",
          items: [
            { id: "home.cells.delete.cells", label: "删除单元格" },
            { id: "home.cells.delete.rows", label: "删除工作表行" },
            { id: "home.cells.delete.cols", label: "删除工作表列" },
          ],
        },
        emit,
      ).element,
    );
    const row2 = document.createElement("div");
    row2.className = "fs-ribbon-stack__row";
    row2.appendChild(
      createToolbarButton({ id: "home.cells.format", tab: TAB, label: "格式", variant: "large" }, emit).element,
    );
    stack.appendChild(row1);
    stack.appendChild(row2);
    content.appendChild(stack);
    inner.appendChild(root);
  }

  // 编辑
  {
    const { root, content } = createRibbonGroup("编辑");
    const row = document.createElement("div");
    row.className = "fs-ribbon-stack__row";
    row.appendChild(createToolbarButton({ id: "home.edit.find", tab: TAB, label: "查找" }, emit).element);
    row.appendChild(createToolbarButton({ id: "home.edit.replace", tab: TAB, label: "替换" }, emit).element);
    row.appendChild(createToolbarButton({ id: "home.edit.clear", tab: TAB, label: "清除" }, emit).element);
    content.appendChild(row);
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
      splitDropdown: true,
      title: "后退 (Ctrl+Z)",
      disabled: true,
    },
    emit,
  );
  back.element.addEventListener("fs-dropdown-toggle", () => {
    emit("home.undo.undoMenu", TAB);
  });
  const forward = createToolbarButton(
    {
      id: "home.undo.forward",
      tab: TAB,
      label: "",
      icon: iconRedo(),
      splitDropdown: true,
      title: "前进 (Ctrl+Y)",
      disabled: true,
    },
    emit,
  );
  forward.element.addEventListener("fs-dropdown-toggle", () => {
    emit("home.undo.redoMenu", TAB);
  });
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
