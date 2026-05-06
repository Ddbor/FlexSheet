import { createToolbarButton, type RibbonEmit } from "../../toolbar/index.js";
import {
  iconFormulaRibbonAutoSum,
  iconFormulaRibbonDateTime,
  iconFormulaRibbonFinancial,
  iconFormulaRibbonLogical,
  iconFormulaRibbonLookup,
  iconFormulaRibbonMath,
  iconFormulaRibbonMore,
  iconFormulaRibbonText,
  iconInsertFunction,
} from "../../toolbar/icons.js";
import { mountAutoSumSubmenu } from "../home-autosum-menu.js";
import { mountFormulaRibbonFnMenu } from "../formula-ribbon-fn-menu.js";
import { createRibbonGroup } from "../ribbon-group.js";
import type { FlexSheetLike, RibbonTabId } from "../ribbon-types.js";

const TAB: RibbonTabId = "formula";

type CategorySpec = {
  readonly menuKey: string;
  readonly label: string;
  readonly icon: SVGSVGElement;
  /** 与 `EXCEL_MS_CATEGORY_BY_NAME` 中文类名一致；`__other__` 表示非主分类 */
  readonly msCategories: readonly string[];
};

const FORMULA_LIBRARY_CATEGORIES: readonly CategorySpec[] = [
  {
    menuKey: "financial",
    label: "财务",
    icon: iconFormulaRibbonFinancial(),
    msCategories: ["财务"],
  },
  {
    menuKey: "logical",
    label: "逻辑",
    icon: iconFormulaRibbonLogical(),
    msCategories: ["逻辑"],
  },
  {
    menuKey: "text",
    label: "文本",
    icon: iconFormulaRibbonText(),
    msCategories: ["文本"],
  },
  {
    menuKey: "datetime",
    label: "日期和时间",
    icon: iconFormulaRibbonDateTime(),
    msCategories: ["日期与时间"],
  },
  {
    menuKey: "lookup",
    label: "查找与引用",
    icon: iconFormulaRibbonLookup(),
    msCategories: ["查找与引用"],
  },
  {
    menuKey: "math",
    label: "数学和三角函数",
    icon: iconFormulaRibbonMath(),
    msCategories: ["数学与三角函数", "数学和三角"],
  },
  {
    menuKey: "other",
    label: "其他函数",
    icon: iconFormulaRibbonMore(),
    msCategories: ["__other__"],
  },
];

export function mountFormulaTab(
  panel: HTMLElement,
  emit: RibbonEmit,
  getFlexSheet: () => FlexSheetLike | undefined,
): void {
  const inner = document.createElement("div");
  inner.className = "fs-ribbon-panel__inner";
  panel.appendChild(inner);

  const { root, content } = createRibbonGroup("函数库");
  content.classList.add("fs-ribbon-formula-function-lib");

  const row = document.createElement("div");
  row.className = "fs-ribbon-formula-function-lib__row";

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

  {
    const sumBtn = createToolbarButton(
      {
        id: "formula.autoSum",
        tab: TAB,
        label: "自动求和",
        icon: iconFormulaRibbonAutoSum(),
        variant: "large",
        splitDropdown: true,
        title: "自动求和（求和）/ 展开其他聚合",
      },
      emit,
    );
    sumBtn.element.id = "fs-ribbon-formula-autosum";
    mountAutoSumSubmenu(sumBtn.element, emit, TAB);
    row.appendChild(sumBtn.element);
  }

  const buildCategoryMenuItems = (spec: CategorySpec) => {
    const fs = getFlexSheet();
    const names = fs?.listFormulaNamesForRibbonCategories?.(spec.msCategories, 48) ?? [];
    const out: { commandId: string; label: string; separatorTop?: boolean; payload?: { name: string } }[] =
      [];
    for (const n of names) {
      out.push({ commandId: "formula.fn.pick", label: n, payload: { name: n } });
    }
    out.push({ commandId: "formula.insertFunction", label: "插入函数…", separatorTop: true });
    return out;
  };

  for (const spec of FORMULA_LIBRARY_CATEGORIES) {
    const b = createToolbarButton(
      {
        id: `formula.cat.${spec.menuKey}`,
        tab: TAB,
        label: spec.label,
        icon: spec.icon,
        variant: "large",
        menuTrigger: true,
        title: spec.label,
      },
      emit,
    );
    b.element.id = `fs-ribbon-formula-cat-${spec.menuKey}`;
    mountFormulaRibbonFnMenu(b.element, emit, TAB, () => buildCategoryMenuItems(spec));
    row.appendChild(b.element);
  }

  content.appendChild(row);
  inner.appendChild(root);
}
