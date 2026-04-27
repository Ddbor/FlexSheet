import { createToolbarButton, createToolbarDropdown, type RibbonEmit } from "../../toolbar/index.js";
import { mountAutoSumSubmenu } from "../home-autosum-menu.js";
import { iconAudit, iconFunction } from "../../toolbar/icons.js";
import { createRibbonGroup } from "../ribbon-group.js";
import type { RibbonTabId } from "../ribbon-types.js";

const TAB: RibbonTabId = "formula";

export function mountFormulaTab(panel: HTMLElement, emit: RibbonEmit): void {
  const inner = document.createElement("div");
  inner.className = "fs-ribbon-panel__inner";
  panel.appendChild(inner);

  {
    const { root, content } = createRibbonGroup("函数库");
    const row = document.createElement("div");
    row.className = "fs-ribbon-stack__row";
    row.appendChild(
      createToolbarDropdown(
        {
          id: "formula.insertFunction",
          tab: TAB,
          label: "插入函数",
          items: [
            { id: "formula.fn.sum", label: "SUM" },
            { id: "formula.fn.average", label: "AVERAGE" },
            { id: "formula.fn.if", label: "IF" },
            { id: "formula.fn.vlookup", label: "VLOOKUP" },
            { id: "formula.fn.more", label: "其他函数…" },
          ],
        },
        emit,
      ).element,
    );
    {
      const asBtn = createToolbarButton(
        {
          id: "formula.autoSum",
          tab: TAB,
          label: "自动求和",
          icon: iconFunction(),
          variant: "large",
          splitDropdown: true,
          title: "自动求和 / 其他函数",
        },
        emit,
      );
      asBtn.element.id = "fs-ribbon-formula-autosum";
      mountAutoSumSubmenu(asBtn.element, emit, TAB);
      row.appendChild(asBtn.element);
    }
    content.appendChild(row);
    inner.appendChild(root);
  }

  {
    const { root, content } = createRibbonGroup("公式审核");
    const row1 = document.createElement("div");
    row1.className = "fs-ribbon-stack__row";
    row1.appendChild(
      createToolbarButton(
        { id: "formula.audit.tracePrecedents", tab: TAB, label: "追踪引用", icon: iconAudit() },
        emit,
      ).element,
    );
    row1.appendChild(
      createToolbarButton(
        { id: "formula.audit.traceDependents", tab: TAB, label: "追踪从属", icon: iconAudit() },
        emit,
      ).element,
    );
    const row2 = document.createElement("div");
    row2.className = "fs-ribbon-stack__row";
    row2.appendChild(
      createToolbarButton({ id: "formula.audit.showFormulas", tab: TAB, label: "显示公式", icon: iconAudit() }, emit)
        .element,
    );
    row2.appendChild(
      createToolbarButton({ id: "formula.audit.errorCheck", tab: TAB, label: "错误检查", icon: iconAudit() }, emit)
        .element,
    );
    content.appendChild(row1);
    content.appendChild(row2);
    inner.appendChild(root);
  }
}
