import { createToolbarButton, type RibbonEmit } from "../../toolbar/index.js";
import { iconInsertTable, iconPivotTableOption } from "../../toolbar/icons.js";
import { createRibbonGroup } from "../ribbon-group.js";
import type { RibbonTabId } from "../ribbon-types.js";

const TAB: RibbonTabId = "insert";

export function mountInsertTab(panel: HTMLElement, emit: RibbonEmit): void {
  const inner = document.createElement("div");
  inner.className = "fs-ribbon-panel__inner";
  panel.appendChild(inner);

  const groups: {
    label: string;
    buttons: {
      id: string;
      label: string;
      icon: () => SVGSVGElement;
      variant?: "default" | "large";
    }[];
  }[] = [
    {
      label: "表格",
      buttons: [
        {
          id: "insert.table",
          label: "表格",
          icon: iconInsertTable,
          variant: "large",
        },
        {
          id: "insert.pivottable.options",
          label: "数据透视表",
          icon: iconPivotTableOption,
          variant: "large",
        },
      ],
    },
  ];

  for (const g of groups) {
    const { root, content } = createRibbonGroup(g.label);
    if (g.label === "表格") {
      content.classList.add("fs-ribbon-insert-pivot");
    }
    const row = document.createElement("div");
    row.className = "fs-ribbon-stack__row";
    for (const b of g.buttons) {
      row.appendChild(
        createToolbarButton(
          { id: b.id, tab: TAB, label: b.label, icon: b.icon(), variant: b.variant },
          emit,
        ).element,
      );
    }
    content.appendChild(row);
    inner.appendChild(root);
  }
}
