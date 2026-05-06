import { createToolbarButton, type RibbonEmit } from "../../toolbar/toolbar-button.js";
import {
  iconPictureChange,
  iconPictureColor,
  iconPictureCorrect,
  iconPictureReset,
  iconPictureTransparency,
} from "../../toolbar/icons.js";
import { mountPictureCorrectionsMenu } from "../picture-corrections-menu.js";
import { mountPictureFormatPlaceholderMenu } from "../picture-format-menus.js";
import { createRibbonGroup } from "../ribbon-group.js";
import type { FlexSheetLike, RibbonTabId } from "../ribbon-types.js";

const TAB: RibbonTabId = "pictureFormat";

/**
 * 「图片格式」上下文选项卡：调整组（更正 / 颜色 / 透明度 为下拉占位；更改图片 / 重置图片 为大按钮占位）。
 */
export function mountPictureFormatTab(
  panel: HTMLElement,
  emit: RibbonEmit,
  getFlexSheet?: () => FlexSheetLike | undefined,
): void {
  const inner = document.createElement("div");
  inner.className = "fs-ribbon-panel__inner";
  panel.appendChild(inner);

  const { root, content } = createRibbonGroup("调整");
  content.classList.add("fs-ribbon-picture-format");

  const row = document.createElement("div");
  row.className = "fs-ribbon-picture-format__row";

  const correct = createToolbarButton(
    {
      id: "pictureFormat.correct",
      tab: TAB,
      label: "更正",
      variant: "large",
      icon: iconPictureCorrect(),
      menuTrigger: true,
      title: "更正",
    },
    emit,
  );
  correct.element.id = "fs-ribbon-picture-format-correct";
  if (getFlexSheet !== undefined) {
    mountPictureCorrectionsMenu(correct.element, getFlexSheet);
  } else {
    mountPictureFormatPlaceholderMenu(correct.element, ["更正（未连接 FlexSheet）"]);
  }

  const color = createToolbarButton(
    {
      id: "pictureFormat.color",
      tab: TAB,
      label: "颜色",
      variant: "large",
      icon: iconPictureColor(),
      menuTrigger: true,
      title: "颜色",
    },
    emit,
  );
  color.element.id = "fs-ribbon-picture-format-color";
  mountPictureFormatPlaceholderMenu(color.element, [
    "颜色饱和度…",
    "色调…",
    "重新着色…",
  ]);

  const transparency = createToolbarButton(
    {
      id: "pictureFormat.transparency",
      tab: TAB,
      label: "透明度",
      variant: "large",
      icon: iconPictureTransparency(),
      menuTrigger: true,
      title: "透明度",
    },
    emit,
  );
  transparency.element.id = "fs-ribbon-picture-format-transparency";
  mountPictureFormatPlaceholderMenu(transparency.element, [
    "预设透明度…",
    "设置透明色…",
  ]);

  const changePicture = createToolbarButton(
    {
      id: "pictureFormat.changePicture",
      tab: TAB,
      label: "更改图片",
      variant: "large",
      icon: iconPictureChange(),
      title: "更改图片",
    },
    emit,
  );
  changePicture.element.id = "fs-ribbon-picture-format-change";

  const resetPicture = createToolbarButton(
    {
      id: "pictureFormat.resetPicture",
      tab: TAB,
      label: "重置图片",
      variant: "large",
      icon: iconPictureReset(),
      title: "重置图片",
    },
    emit,
  );
  resetPicture.element.id = "fs-ribbon-picture-format-reset";

  row.appendChild(correct.element);
  row.appendChild(color.element);
  row.appendChild(transparency.element);
  row.appendChild(changePicture.element);
  row.appendChild(resetPicture.element);
  content.appendChild(row);
  inner.appendChild(root);
}
