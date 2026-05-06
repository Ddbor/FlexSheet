import { createToolbarButton, type RibbonEmit } from "../../toolbar/index.js";
import { iconInsertPicture, iconInsertTable } from "../../toolbar/icons.js";
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
      /** 在该按钮前绘制竖向分割线（同组内） */
      separatorBefore?: boolean;
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
          id: "insert.picture",
          label: "图片",
          icon: iconInsertPicture,
          variant: "large",
          separatorBefore: true,
        },
        // {
        //   id: "insert.pivottable.options",
        //   label: "数据透视表",
        //   icon: iconPivotTableOption,
        //   variant: "large",
        // },
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
      if (b.separatorBefore === true) {
        const sep = document.createElement("div");
        sep.className = "fs-ribbon-insert__sep";
        sep.setAttribute("aria-hidden", "true");
        row.appendChild(sep);
      }
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
