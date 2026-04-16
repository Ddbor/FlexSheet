import type { RibbonTabId } from "./ribbon-types.js";
import type { RibbonEmit } from "../toolbar/toolbar-button.js";
import {
  closeAllRibbonPopups,
  clearToolbarDropdownMenuPosition,
  syncToolbarDropdownMenuPosition,
} from "../toolbar/toolbar-dropdown.js";

function secHead(text: string): HTMLDivElement {
  const el = document.createElement("div");
  el.className = "fs-bd-menu__cell-sec-head";
  el.textContent = text;
  return el;
}

function swatchBtn(
  commandId: string,
  label: string,
  previewStyle: { readonly bg: string; readonly color: string; readonly fontWeight?: string; readonly fontStyle?: string; readonly border?: string },
): HTMLButtonElement {
  const b = document.createElement("button");
  b.type = "button";
  b.className = "fs-bd-menu__cell-swatch";
  b.dataset.commandId = commandId;
  b.setAttribute("aria-label", label);
  const pv = document.createElement("span");
  pv.className = "fs-bd-menu__cell-swatch-preview";
  pv.style.background = previewStyle.bg;
  pv.style.color = previewStyle.color;
  if (previewStyle.fontWeight !== undefined) {
    pv.style.fontWeight = previewStyle.fontWeight;
  }
  if (previewStyle.fontStyle !== undefined) {
    pv.style.fontStyle = previewStyle.fontStyle;
  }
  if (previewStyle.border !== undefined) {
    pv.style.border = previewStyle.border;
  }
  pv.textContent = "Aa";
  const lb = document.createElement("span");
  lb.className = "fs-bd-menu__cell-swatch-label";
  lb.textContent = label;
  b.appendChild(pv);
  b.appendChild(lb);
  return b;
}

function rowItem(
  commandId: string,
  label: string,
  previewStyle: { readonly bg: string; readonly color: string; readonly fontWeight?: string; readonly fontStyle?: string; readonly textDecoration?: string },
): HTMLButtonElement {
  const b = document.createElement("button");
  b.type = "button";
  b.className = "fs-bd-menu__cell-row";
  b.dataset.commandId = commandId;
  const pv = document.createElement("span");
  pv.className = "fs-bd-menu__cell-row-preview";
  pv.style.background = previewStyle.bg;
  pv.style.color = previewStyle.color;
  if (previewStyle.fontWeight !== undefined) {
    pv.style.fontWeight = previewStyle.fontWeight;
  }
  if (previewStyle.fontStyle !== undefined) {
    pv.style.fontStyle = previewStyle.fontStyle;
  }
  if (previewStyle.textDecoration !== undefined) {
    pv.style.textDecoration = previewStyle.textDecoration;
  }
  pv.textContent = "Aa";
  const lb = document.createElement("span");
  lb.className = "fs-bd-menu__cell-row-label";
  lb.textContent = label;
  b.appendChild(pv);
  b.appendChild(lb);
  return b;
}

const THEME_ACCENTS = ["#4472C4", "#ED7D31", "#A5A5A5", "#FFC000", "#5B9BD5", "#70AD47"] as const;
const THEME_TINTS = [0.2, 0.4, 0.6, 1] as const;

function blendHexWithWhite(hex: string, t: number): string {
  const r = parseInt(hex.slice(1, 3), 16);
  const g = parseInt(hex.slice(3, 5), 16);
  const b = parseInt(hex.slice(5, 7), 16);
  const mix = (c: number): number => Math.round(255 * (1 - t) + c * t);
  const rr = mix(r);
  const gg = mix(g);
  const bb = mix(b);
  return `rgb(${rr}, ${gg}, ${bb})`;
}

function iconNewStyle(): SVGSVGElement {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("width", "16");
  svg.setAttribute("height", "16");
  svg.setAttribute("viewBox", "0 0 16 16");
  svg.setAttribute("fill", "none");
  svg.setAttribute("stroke", "currentColor");
  svg.setAttribute("stroke-width", "1.2");
  const g = document.createElementNS("http://www.w3.org/2000/svg", "g");
  const r = document.createElementNS("http://www.w3.org/2000/svg", "rect");
  r.setAttribute("x", "2");
  r.setAttribute("y", "2");
  r.setAttribute("width", "10");
  r.setAttribute("height", "10");
  r.setAttribute("rx", "1");
  const p = document.createElementNS("http://www.w3.org/2000/svg", "path");
  p.setAttribute("d", "M10 6l4-4M10 2h4v4");
  g.appendChild(r);
  g.appendChild(p);
  svg.appendChild(g);
  return svg;
}

/** 「开始 → 单元格样式」：与 Excel 结构一致的分段样式库。 */
export function mountCellStyleRibbonMenu(
  anchor: HTMLButtonElement,
  emit: RibbonEmit,
  tab: RibbonTabId,
): void {
  const menu = document.createElement("div");
  menu.className = "fs-bd-menu fs-bd-menu--cell-styles";
  menu.hidden = true;
  menu.setAttribute("role", "menu");
  menu.setAttribute("data-fs-floating-menu", "");
  menu.setAttribute("data-fs-menu-anchor-id", anchor.id);

  const close = (): void => {
    menu.hidden = true;
    clearToolbarDropdownMenuPosition(menu);
    anchor.setAttribute("aria-expanded", "false");
  };

  const wireCmd = (root: HTMLElement): void => {
    root.addEventListener("click", (ev) => {
      const t = ev.target;
      if (!(t instanceof HTMLElement)) {
        return;
      }
      const btn = t.closest("button[data-command-id]");
      if (!(btn instanceof HTMLButtonElement)) {
        return;
      }
      const id = btn.dataset.commandId;
      if (id === undefined) {
        return;
      }
      ev.stopPropagation();
      emit(id, tab);
      close();
    });
  };

  // 好、差和适中
  menu.appendChild(secHead("好、差和适中"));
  const row1 = document.createElement("div");
  row1.className = "fs-bd-menu__cell-swatch-row";
  row1.appendChild(
    swatchBtn("home.style.cell.normal", "常规", {
      bg: "#ffffff",
      color: "#000000",
      border: "1px solid #bfbfbf",
    }),
  );
  row1.appendChild(
    swatchBtn("home.style.cell.bad", "差", {
      bg: "#ffc7ce",
      color: "#9c0006",
      fontWeight: "700",
      border: "1px solid #bfbfbf",
    }),
  );
  row1.appendChild(
    swatchBtn("home.style.cell.good", "好", {
      bg: "#c6efce",
      color: "#006100",
      fontWeight: "700",
      border: "1px solid #bfbfbf",
    }),
  );
  row1.appendChild(
    swatchBtn("home.style.cell.neutral", "适中", {
      bg: "#ffeb9c",
      color: "#9c6500",
      border: "1px solid #bfbfbf",
    }),
  );
  menu.appendChild(row1);

  // 数据和模型
  menu.appendChild(secHead("数据和模型"));
  const dataCol = document.createElement("div");
  dataCol.className = "fs-bd-menu__cell-col";
  dataCol.appendChild(
    rowItem("home.style.cell.calculation", "计算", {
      bg: "#f2f2f2",
      color: "#c65911",
      fontWeight: "700",
    }),
  );
  dataCol.appendChild(
    rowItem("home.style.cell.checkCell", "检查单元格", {
      bg: "#595959",
      color: "#ffffff",
      fontWeight: "700",
    }),
  );
  dataCol.appendChild(
    rowItem("home.style.cell.explanatory", "解释性文本", {
      bg: "#ffffff",
      color: "#7f7f7f",
      fontStyle: "italic",
    }),
  );
  dataCol.appendChild(
    rowItem("home.style.cell.warningText", "警告文本", {
      bg: "#ffffff",
      color: "#ff0000",
    }),
  );
  dataCol.appendChild(
    rowItem("home.style.cell.linkedCell", "链接单元格", {
      bg: "#ffffff",
      color: "#ff6600",
      textDecoration: "underline double",
    }),
  );
  dataCol.appendChild(
    rowItem("home.style.cell.output", "输出", {
      bg: "#f2f2f2",
      color: "#000000",
      fontWeight: "700",
    }),
  );
  dataCol.appendChild(
    rowItem("home.style.cell.input", "输入", {
      bg: "#fde9d9",
      color: "#974706",
    }),
  );
  dataCol.appendChild(
    rowItem("home.style.cell.note", "注释", {
      bg: "#fff2cc",
      color: "#000000",
    }),
  );
  menu.appendChild(dataCol);

  // 标题
  menu.appendChild(secHead("标题"));
  const titleCol = document.createElement("div");
  titleCol.className = "fs-bd-menu__cell-col";
  titleCol.appendChild(
    rowItem("home.style.cell.title", "标题", {
      bg: "#ffffff",
      color: "#1f497d",
      fontWeight: "700",
    }),
  );
  titleCol.appendChild(
    rowItem("home.style.cell.heading1", "标题 1", {
      bg: "#ffffff",
      color: "#1f497d",
      fontWeight: "700",
    }),
  );
  titleCol.appendChild(
    rowItem("home.style.cell.heading2", "标题 2", {
      bg: "#ffffff",
      color: "#1f497d",
      fontWeight: "700",
    }),
  );
  titleCol.appendChild(
    rowItem("home.style.cell.heading3", "标题 3", {
      bg: "#ffffff",
      color: "#1f497d",
      fontWeight: "700",
    }),
  );
  titleCol.appendChild(
    rowItem("home.style.cell.heading4", "标题 4", {
      bg: "#ffffff",
      color: "#1f497d",
      fontWeight: "700",
    }),
  );
  titleCol.appendChild(
    rowItem("home.style.cell.total", "汇总", {
      bg: "#ffffff",
      color: "#000000",
      fontWeight: "700",
    }),
  );
  menu.appendChild(titleCol);

  // 主题单元格样式
  menu.appendChild(secHead("主题单元格样式"));
  const themeWrap = document.createElement("div");
  themeWrap.className = "fs-bd-menu__cell-theme";
  const headRow = document.createElement("div");
  headRow.className = "fs-bd-menu__cell-theme-head";
  headRow.appendChild(document.createElement("div"));
  for (let c = 0; c < 6; c++) {
    const h = document.createElement("div");
    h.className = "fs-bd-menu__cell-theme-h";
    h.textContent = `着色 ${c + 1}`;
    headRow.appendChild(h);
  }
  themeWrap.appendChild(headRow);
  for (let r = 0; r < 4; r++) {
    const tr = document.createElement("div");
    tr.className = "fs-bd-menu__cell-theme-row";
    const rl = document.createElement("div");
    rl.className = "fs-bd-menu__cell-theme-rlabel";
    rl.textContent = r === 3 ? "100%" : `${[20, 40, 60][r]}%`;
    tr.appendChild(rl);
    for (let c = 0; c < 6; c++) {
      const cell = document.createElement("button");
      cell.type = "button";
      cell.className = "fs-bd-menu__cell-theme-cell";
      cell.dataset.commandId = `home.style.cell.theme.r${r}.c${c}`;
      const t = THEME_TINTS[r];
      const bg = blendHexWithWhite(THEME_ACCENTS[c], t);
      cell.style.background = bg;
      const solid = r === 3;
      cell.style.color = solid ? "#ffffff" : "#000000";
      cell.style.fontWeight = solid ? "700" : "400";
      cell.textContent = "Aa";
      cell.setAttribute(
        "aria-label",
        `着色 ${c + 1} ${r === 3 ? "100%" : `${[20, 40, 60][r]}%`}`,
      );
      tr.appendChild(cell);
    }
    themeWrap.appendChild(tr);
  }
  menu.appendChild(themeWrap);

  // 数字格式
  menu.appendChild(secHead("数字格式"));
  const numCol = document.createElement("div");
  numCol.className = "fs-bd-menu__cell-col";
  numCol.appendChild(rowItem("home.style.cell.num.percent", "百分比", { bg: "#ffffff", color: "#000000" }));
  numCol.appendChild(rowItem("home.style.cell.num.currency", "货币", { bg: "#ffffff", color: "#000000" }));
  numCol.appendChild(rowItem("home.style.cell.num.currency0", "货币[0]", { bg: "#ffffff", color: "#000000" }));
  numCol.appendChild(rowItem("home.style.cell.num.comma", "千位分隔", { bg: "#ffffff", color: "#000000" }));
  numCol.appendChild(rowItem("home.style.cell.num.comma0", "千位分隔[0]", { bg: "#ffffff", color: "#000000" }));
  menu.appendChild(numCol);

  const sep = document.createElement("div");
  sep.className = "fs-bd-menu__sep";
  sep.setAttribute("role", "separator");
  sep.setAttribute("aria-hidden", "true");
  menu.appendChild(sep);

  const foot1 = document.createElement("button");
  foot1.type = "button";
  foot1.className = "fs-bd-menu__item";
  foot1.setAttribute("role", "menuitem");
  foot1.dataset.commandId = "home.style.cell.newStyle";
  const ic1 = document.createElement("span");
  ic1.className = "fs-bd-menu__icon";
  ic1.appendChild(iconNewStyle());
  const lb1 = document.createElement("span");
  lb1.className = "fs-bd-menu__label";
  lb1.textContent = "新建单元格样式...";
  foot1.appendChild(ic1);
  foot1.appendChild(lb1);

  menu.appendChild(foot1);

  wireCmd(menu);

  const ribbonRoot = anchor.closest(".fs-ribbon");
  (ribbonRoot ?? document.body).appendChild(menu);
  anchor.setAttribute("aria-haspopup", "menu");
  anchor.setAttribute("aria-expanded", "false");

  anchor.addEventListener("fs-dropdown-toggle", (ev) => {
    ev.stopPropagation();
    if (!menu.hidden) {
      close();
      return;
    }
    closeAllRibbonPopups();
    menu.hidden = false;
    syncToolbarDropdownMenuPosition(anchor, menu);
    anchor.setAttribute("aria-expanded", "true");
  });
}
