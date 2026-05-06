import type { RibbonEmit } from "../../toolbar/toolbar-button.js";
import { createRibbonGroup } from "../ribbon-group.js";

/**
 * 「图片格式」上下文选项卡（浮动图片选中时显示）；具体命令后续再接。
 */
export function mountPictureFormatTab(panel: HTMLElement, _emit: RibbonEmit): void {
  const inner = document.createElement("div");
  inner.className = "fs-ribbon-panel__inner";
  panel.appendChild(inner);

  const { root, content } = createRibbonGroup("图片");
  inner.appendChild(root);
  content.setAttribute("aria-label", "图片格式工具占位");
}
