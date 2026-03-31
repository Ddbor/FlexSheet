/**
 * Ribbon 分组容器：底栏标题 + 顶区工具。
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
