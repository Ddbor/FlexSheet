import { createToolbarButton, type RibbonEmit } from "../../toolbar/index.js";
import { iconChart, iconImage, iconShapes, iconSparkline, iconTable } from "../../toolbar/icons.js";
import { createRibbonGroup } from "../ribbon-group.js";
import type { RibbonTabId } from "../ribbon-types.js";

const TAB: RibbonTabId = "insert";

export function mountInsertTab(panel: HTMLElement, emit: RibbonEmit): void {
  const inner = document.createElement("div");
  inner.className = "fs-ribbon-panel__inner";
  panel.appendChild(inner);

  const groups: {
    label: string;
    buttons: { id: string; label: string; icon: () => SVGSVGElement; variant?: "default" | "large" }[];
  }[] = [
    {
      label: "表格",
      buttons: [{ id: "insert.table", label: "表格", icon: iconTable, variant: "large" }],
    },
    {
      label: "插图",
      buttons: [
        { id: "insert.picture", label: "图片", icon: iconImage, variant: "large" },
        { id: "insert.icons", label: "图标", icon: iconShapes, variant: "large" },
      ],
    },
    {
      label: "图表",
      buttons: [
        { id: "insert.chart.column", label: "柱形图", icon: iconChart },
        { id: "insert.chart.line", label: "折线图", icon: iconChart },
        { id: "insert.chart.pie", label: "饼图", icon: iconChart },
      ],
    },
    {
      label: "迷你图",
      buttons: [
        { id: "insert.sparkline.line", label: "折线", icon: iconSparkline },
        { id: "insert.sparkline.column", label: "柱状", icon: iconSparkline },
        { id: "insert.sparkline.winloss", label: "盈亏", icon: iconSparkline },
      ],
    },
  ];

  for (const g of groups) {
    const { root, content } = createRibbonGroup(g.label);
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
