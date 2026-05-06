import { createToolbarButton, type RibbonEmit } from "../../toolbar/index.js";
import { iconInsertFunction } from "../../toolbar/icons.js";
import { createRibbonGroup } from "../ribbon-group.js";
import type { RibbonTabId } from "../ribbon-types.js";

const TAB: RibbonTabId = "formula";

export function mountFormulaTab(panel: HTMLElement, emit: RibbonEmit): void {
  const inner = document.createElement("div");
  inner.className = "fs-ribbon-panel__inner";
  panel.appendChild(inner);

  const { root, content } = createRibbonGroup("函数库");
  content.classList.add("fs-ribbon-formula-function-lib");
  const row = document.createElement("div");
  row.className = "fs-ribbon-stack__row";
  row.appendChild(
    createToolbarButton(
      {
        id: "formula.insertFunction",
        tab: TAB,
        label: "插入函数",
        icon: iconInsertFunction(),
        variant: "large",
      },
      emit,
    ).element,
  );
  content.appendChild(row);
  inner.appendChild(root);
}
