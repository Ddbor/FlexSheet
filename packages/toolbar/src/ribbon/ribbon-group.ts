/**
 * Ribbon 分组容器：顶区工具；底栏分组名节点保留但不展示（见 FlexSheetRibbon.css）。
 */

export function createRibbonGroup(label: string): {
  readonly root: HTMLElement;
  readonly content: HTMLElement;
} {
  const root = document.createElement("div");
  root.className = "fs-ribbon-group";

  const content = document.createElement("div");
  content.className = "fs-ribbon-group__content";

  const foot = document.createElement("div");
  foot.className = "fs-ribbon-group__label";
  foot.textContent = label;

  root.appendChild(content);
  root.appendChild(foot);
  return { root, content };
}
