import { createToolbarButton, createToolbarDropdown, type RibbonEmit } from "../../toolbar/index.js";
import { iconBackground, iconMargins, iconOrientation, iconPrintTitles } from "../../toolbar/icons.js";
import { createRibbonGroup } from "../ribbon-group.js";
import type { RibbonTabId } from "../ribbon-types.js";

const TAB: RibbonTabId = "pageLayout";

export function mountPageLayoutTab(panel: HTMLElement, emit: RibbonEmit): void {
  const inner = document.createElement("div");
  inner.className = "fs-ribbon-panel__inner";
  panel.appendChild(inner);

  {
    const { root, content } = createRibbonGroup("页面设置");
    const row = document.createElement("div");
    row.className = "fs-ribbon-stack__row";
    row.appendChild(
      createToolbarButton(
        { id: "pageLayout.orientation.portrait", tab: TAB, label: "纵向", icon: iconOrientation(), variant: "large" },
        emit,
      ).element,
    );
    row.appendChild(
      createToolbarButton(
        { id: "pageLayout.orientation.landscape", tab: TAB, label: "横向", icon: iconOrientation(), variant: "large" },
        emit,
      ).element,
    );
    content.appendChild(row);
    inner.appendChild(root);
  }

  {
    const { root, content } = createRibbonGroup("页边距");
    content.appendChild(
      createToolbarDropdown(
        {
          id: "pageLayout.margins",
          tab: TAB,
          label: "常规",
          items: [
            { id: "pageLayout.margins.normal", label: "常规" },
            { id: "pageLayout.margins.narrow", label: "窄" },
            { id: "pageLayout.margins.wide", label: "宽" },
            { id: "pageLayout.margins.custom", label: "自定义边距…" },
          ],
        },
        emit,
      ).element,
    );
    const row = document.createElement("div");
    row.className = "fs-ribbon-stack__row";
    row.appendChild(createToolbarButton({ id: "pageLayout.margins.icon", tab: TAB, label: "边距", icon: iconMargins() }, emit).element);
    content.appendChild(row);
    inner.appendChild(root);
  }

  {
    const { root, content } = createRibbonGroup("页面背景");
    content.appendChild(createToolbarButton({ id: "pageLayout.background", tab: TAB, label: "背景", icon: iconBackground(), variant: "large" }, emit).element);
    inner.appendChild(root);
  }

  {
    const { root, content } = createRibbonGroup("打印标题");
    content.appendChild(
      createToolbarButton({ id: "pageLayout.printTitles", tab: TAB, label: "打印标题", icon: iconPrintTitles() }, emit).element,
    );
    inner.appendChild(root);
  }
}
