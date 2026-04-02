/**
 * Office 风格「文件」Backstage：左侧导航 + 右侧内容（当前仅 UI）。
 */

import "./ribbon-backstage.css";

export type RibbonBackstageSectionId = "open" | "save" | "import" | "export" | "print" | "info";

export interface RibbonBackstageHandles {
  readonly root: HTMLElement;
  showSection(id: RibbonBackstageSectionId): void;
  applyThemeMode(mode: "light" | "dark"): void;
}

const SECTION_ORDER: RibbonBackstageSectionId[] = ["open", "save", "import", "export", "print", "info"];

const SECTION_LABEL: Record<RibbonBackstageSectionId, string> = {
  open: "打开",
  save: "保存",
  import: "导入",
  export: "导出",
  print: "打印",
  info: "信息",
};

/** 侧栏在哪些项之后插入分隔线（与参考图一致） */
const DIVIDER_AFTER: Partial<Record<RibbonBackstageSectionId, boolean>> = {
  save: true,
  export: true,
  print: true,
};

function iconBack(): SVGSVGElement {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("width", "16");
  svg.setAttribute("height", "16");
  svg.setAttribute("viewBox", "0 0 16 16");
  svg.setAttribute("aria-hidden", "true");
  const p = document.createElementNS("http://www.w3.org/2000/svg", "path");
  p.setAttribute(
    "d",
    "M10.5 3.5L6 8l4.5 4.5-1 1L4 8 9.5 2.5l1 1z",
  );
  p.setAttribute("fill", "currentColor");
  svg.appendChild(p);
  return svg;
}

function iconDocument(): SVGSVGElement {
  const svg = document.createElementNS("http://www.w3.org/2000/svg", "svg");
  svg.setAttribute("width", "48");
  svg.setAttribute("height", "48");
  svg.setAttribute("viewBox", "0 0 48 48");
  svg.setAttribute("aria-hidden", "true");
  const r = document.createElementNS("http://www.w3.org/2000/svg", "rect");
  r.setAttribute("x", "10");
  r.setAttribute("y", "6");
  r.setAttribute("width", "28");
  r.setAttribute("height", "36");
  r.setAttribute("rx", "2");
  r.setAttribute("fill", "none");
  r.setAttribute("stroke", "currentColor");
  r.setAttribute("stroke-width", "2");
  svg.appendChild(r);
  const line = document.createElementNS("http://www.w3.org/2000/svg", "path");
  line.setAttribute("d", "M14 16h20M14 22h16M14 28h20");
  line.setAttribute("stroke", "currentColor");
  line.setAttribute("stroke-width", "1.5");
  svg.appendChild(line);
  return svg;
}

function createCheckboxRow(label: string, checked: boolean): HTMLElement {
  const row = document.createElement("label");
  row.className = "fs-backstage__check-row";
  const input = document.createElement("input");
  input.type = "checkbox";
  input.checked = checked;
  const span = document.createElement("span");
  span.textContent = label;
  row.appendChild(input);
  row.appendChild(span);
  return row;
}

function mountOpenPanel(container: HTMLElement): void {
  container.innerHTML = "";
  const title = document.createElement("h1");
  title.className = "fs-backstage__page-title";
  title.textContent = "打开";

  const sub = document.createElement("p");
  sub.className = "fs-backstage__page-sub";
  sub.textContent = "SpreadJS 文件 (.sjs)";

  const groups = document.createElement("div");
  groups.className = "fs-backstage__open-groups";

  const g1 = document.createElement("div");
  g1.className = "fs-backstage__open-group";
  g1.appendChild(createCheckboxRow("增量加载", false));
  g1.appendChild(createCheckboxRow("懒加载", false));

  const g2 = document.createElement("div");
  g2.className = "fs-backstage__open-group";
  g2.appendChild(createCheckboxRow("包含样式", true));
  g2.appendChild(createCheckboxRow("包含公式", true));
  g2.appendChild(createCheckboxRow("导入后自动计算", false));

  const g3 = document.createElement("div");
  g3.className = "fs-backstage__open-group";
  g3.appendChild(createCheckboxRow("导入未使用的命名样式", true));
  g3.appendChild(createCheckboxRow("增量计算", true));
  g3.appendChild(createCheckboxRow("按需计算", true));
  g3.appendChild(createCheckboxRow("动态引用", false));

  groups.appendChild(g1);
  groups.appendChild(g2);
  groups.appendChild(g3);

  const card = document.createElement("button");
  card.type = "button";
  card.className = "fs-backstage__open-card";
  const iconWrap = document.createElement("div");
  iconWrap.className = "fs-backstage__open-card-icon";
  iconWrap.appendChild(iconDocument());
  const cardLabel = document.createElement("span");
  cardLabel.textContent = "打开 SJS 文件";
  card.appendChild(iconWrap);
  card.appendChild(cardLabel);

  container.appendChild(title);
  container.appendChild(sub);
  container.appendChild(groups);
  container.appendChild(card);
}

function mountPlaceholderPanel(container: HTMLElement, section: RibbonBackstageSectionId): void {
  container.innerHTML = "";
  const title = document.createElement("h1");
  title.className = "fs-backstage__page-title";
  title.textContent = SECTION_LABEL[section];
  const hint = document.createElement("p");
  hint.className = "fs-backstage__placeholder";
  hint.textContent = "此功能内容稍后补充。";
  container.appendChild(title);
  container.appendChild(hint);
}

export function createRibbonBackstage(onClose: () => void): RibbonBackstageHandles {
  const root = document.createElement("div");
  root.className = "fs-ribbon-backstage";
  root.setAttribute("role", "dialog");
  root.setAttribute("aria-modal", "true");
  root.setAttribute("aria-label", "文件");

  const sidebar = document.createElement("aside");
  sidebar.className = "fs-ribbon-backstage__sidebar";

  const backWrap = document.createElement("div");
  backWrap.className = "fs-ribbon-backstage__back-wrap";
  const backBtn = document.createElement("button");
  backBtn.type = "button";
  backBtn.className = "fs-ribbon-backstage__back";
  backBtn.setAttribute("aria-label", "返回");
  backBtn.title = "返回";
  backBtn.appendChild(iconBack());
  backBtn.addEventListener("click", () => {
    onClose();
  });
  backWrap.appendChild(backBtn);
  sidebar.appendChild(backWrap);

  const nav = document.createElement("nav");
  nav.className = "fs-ribbon-backstage__nav";
  nav.setAttribute("aria-label", "文件菜单");

  const navButtons = new Map<RibbonBackstageSectionId, HTMLButtonElement>();
  const contentArea = document.createElement("div");
  contentArea.className = "fs-ribbon-backstage__main";

  const panels = new Map<RibbonBackstageSectionId, HTMLElement>();

  for (const id of SECTION_ORDER) {
    const itemWrap = document.createElement("div");
    itemWrap.className = "fs-ribbon-backstage__nav-item-wrap";

    const btn = document.createElement("button");
    btn.type = "button";
    btn.className = "fs-ribbon-backstage__nav-item";
    btn.textContent = SECTION_LABEL[id];
    btn.dataset.section = id;
    navButtons.set(id, btn);
    itemWrap.appendChild(btn);
    nav.appendChild(itemWrap);

    if (DIVIDER_AFTER[id] === true) {
      const hr = document.createElement("div");
      hr.className = "fs-ribbon-backstage__nav-divider";
      hr.setAttribute("role", "separator");
      nav.appendChild(hr);
    }

    const panel = document.createElement("div");
    panel.className = "fs-ribbon-backstage__panel";
    panel.hidden = id !== "open";
    panel.dataset.section = id;
    panels.set(id, panel);
    contentArea.appendChild(panel);
  }

  function showSection(id: RibbonBackstageSectionId): void {
    for (const [sid, panel] of panels) {
      panel.hidden = sid !== id;
    }
    for (const [sid, b] of navButtons) {
      b.classList.toggle("fs-ribbon-backstage__nav-item--active", sid === id);
    }
    if (id === "open") {
      mountOpenPanel(panels.get("open")!);
    } else {
      mountPlaceholderPanel(panels.get(id)!, id);
    }
  }

  for (const id of SECTION_ORDER) {
    const b = navButtons.get(id)!;
    b.addEventListener("click", () => {
      showSection(id);
    });
  }

  mountOpenPanel(panels.get("open")!);
  navButtons.get("open")!.classList.add("fs-ribbon-backstage__nav-item--active");

  sidebar.appendChild(nav);
  root.appendChild(sidebar);
  root.appendChild(contentArea);

  root.addEventListener("keydown", (ev: KeyboardEvent) => {
    if (ev.key === "Escape") {
      ev.preventDefault();
      onClose();
    }
  });

  return {
    root,
    showSection,
    applyThemeMode(mode: "light" | "dark"): void {
      root.dataset.theme = mode;
    },
  };
}
