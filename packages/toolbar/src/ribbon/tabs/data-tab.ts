import { createToolbarButton, createToolbarDropdown, type RibbonEmit } from "../../toolbar/index.js";
import { iconDataTools, iconFilter, iconPivotTableOption, iconSlicer, iconSort } from "../../toolbar/icons.js";
import { createRibbonGroup } from "../ribbon-group.js";
import type { RibbonTabId } from "../ribbon-types.js";

const TAB: RibbonTabId = "data";

export function mountDataTab(panel: HTMLElement, emit: RibbonEmit): void {
  const inner = document.createElement("div");
  inner.className = "fs-ribbon-panel__inner";
  panel.appendChild(inner);

  {
    const { root, content } = createRibbonGroup("排序和筛选");
    const row = document.createElement("div");
    row.className = "fs-ribbon-stack__row";
    row.appendChild(createToolbarButton({ id: "data.sort.asc", tab: TAB, label: "升序", icon: iconSort() }, emit).element);
    row.appendChild(createToolbarButton({ id: "data.sort.desc", tab: TAB, label: "降序", icon: iconSort() }, emit).element);
    row.appendChild(createToolbarButton({ id: "data.sort.custom", tab: TAB, label: "自定义排序", icon: iconSort() }, emit).element);
    row.appendChild(createToolbarButton({ id: "data.filter.toggle", tab: TAB, label: "筛选", icon: iconFilter() }, emit).element);
    content.appendChild(row);
    inner.appendChild(root);
  }

  {
    const { root, content } = createRibbonGroup("筛选器");
    content.appendChild(
      createToolbarButton(
        { id: "data.slicer.insert", tab: TAB, label: "插入切片器", icon: iconSlicer(), variant: "large" },
        emit,
      ).element,
    );
    inner.appendChild(root);
  }

  {
    const { root, content } = createRibbonGroup("数据透视表");
    content.appendChild(
      createToolbarButton(
        {
          id: "data.pivot.fields",
          tab: TAB,
          label: "字段列表",
          icon: iconPivotTableOption(),
          variant: "large",
        },
        emit,
      ).element,
    );
    inner.appendChild(root);
  }

  {
    const { root, content } = createRibbonGroup("数据工具");
    const row = document.createElement("div");
    row.className = "fs-ribbon-stack__row";
    row.appendChild(
      createToolbarDropdown(
        {
          id: "data.tools.textToColumns",
          tab: TAB,
          label: "分列",
          items: [
            { id: "data.tools.textToColumns.delimited", label: "分隔符号" },
            { id: "data.tools.textToColumns.fixedWidth", label: "固定宽度" },
          ],
        },
        emit,
      ).element,
    );
    row.appendChild(createToolbarButton({ id: "data.tools.removeDuplicates", tab: TAB, label: "删除重复项", icon: iconDataTools() }, emit).element);
    row.appendChild(createToolbarButton({ id: "data.tools.dataValidation", tab: TAB, label: "数据验证", icon: iconDataTools() }, emit).element);
    content.appendChild(row);
    inner.appendChild(root);
  }
}
