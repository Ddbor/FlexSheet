import { createToolbarButton, type RibbonEmit } from "../../toolbar/toolbar-button.js";
import {
  iconPictureChange,
  iconPictureColor,
  iconPictureCorrect,
  iconPictureReset,
  iconPictureRotate,
  iconPictureTransparency,
} from "../../toolbar/icons.js";
import { mountPictureColorMenu } from "../picture-color-menu.js";
import { mountPictureCorrectionsMenu } from "../picture-corrections-menu.js";
import { mountPictureFormatPlaceholderMenu } from "../picture-format-menus.js";
import { mountPictureRotateMenu } from "../picture-rotate-menu.js";
import { mountPictureTransparencyMenu } from "../picture-transparency-menu.js";
import { createRibbonGroup } from "../ribbon-group.js";
import type { FlexSheetLike, RibbonTabId } from "../ribbon-types.js";

const TAB: RibbonTabId = "pictureFormat";

/**
 * 「图片格式」上下文选项卡：调整组 + 右侧大小（高度/宽度，无图标）。
 * @returns 卸载时取消「大小」区对 FlexSheet 的订阅。
 */
export function mountPictureFormatTab(
  panel: HTMLElement,
  emit: RibbonEmit,
  getFlexSheet?: () => FlexSheetLike | undefined,
): () => void {
  const cleanups: Array<() => void> = [];

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
  if (getFlexSheet !== undefined) {
    mountPictureColorMenu(color.element, getFlexSheet);
  } else {
    mountPictureFormatPlaceholderMenu(color.element, [
      "颜色饱和度…",
      "色调…",
      "重新着色…",
    ]);
  }

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
  if (getFlexSheet !== undefined) {
    mountPictureTransparencyMenu(transparency.element, getFlexSheet);
  } else {
    mountPictureFormatPlaceholderMenu(transparency.element, [
      "预设透明度…",
      "设置透明色…",
    ]);
  }

  const rotate = createToolbarButton(
    {
      id: "pictureFormat.rotate",
      tab: TAB,
      label: "旋转",
      variant: "large",
      icon: iconPictureRotate(),
      menuTrigger: true,
      title: "旋转",
    },
    emit,
  );
  rotate.element.id = "fs-ribbon-picture-format-rotate";
  if (getFlexSheet !== undefined) {
    mountPictureRotateMenu(rotate.element, getFlexSheet);
  } else {
    mountPictureFormatPlaceholderMenu(rotate.element, ["向右旋转 90°", "向左旋转 90°"]);
  }

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
  row.appendChild(rotate.element);
  row.appendChild(changePicture.element);
  row.appendChild(resetPicture.element);
  content.appendChild(row);
  inner.appendChild(root);

  if (getFlexSheet !== undefined) {
    const getFs = getFlexSheet;
    const { root: sizeRoot, content: sizeContent } = createRibbonGroup("大小");
    sizeContent.classList.add("fs-ribbon-picture-size");

    let syncingPicSize = false;

    const mkRow = (labelText: string): { row: HTMLDivElement; input: HTMLInputElement } => {
      const r = document.createElement("div");
      r.className = "fs-ribbon-picture-size__row";
      const lab = document.createElement("label");
      lab.className = "fs-ribbon-picture-size__label";
      lab.textContent = labelText;
      const inp = document.createElement("input");
      inp.type = "number";
      inp.className = "fs-ribbon-picture-size__input";
      inp.min = "24";
      inp.step = "1";
      inp.setAttribute("aria-label", labelText);
      r.appendChild(lab);
      r.appendChild(inp);
      return { row: r, input: inp };
    };

    const hRow = mkRow("高度");
    const wRow = mkRow("宽度");
    sizeContent.appendChild(hRow.row);
    sizeContent.appendChild(wRow.row);

    const MIN_PIC_PX = 24;

    const refreshPicSizeInputs = (): void => {
      const fs = getFs();
      if (fs === undefined) {
        return;
      }
      const lay = fs.getSelectedFloatingPictureLayoutPx?.() ?? null;
      syncingPicSize = true;
      try {
        if (lay === null) {
          hRow.input.value = "";
          wRow.input.value = "";
          hRow.input.disabled = true;
          wRow.input.disabled = true;
          return;
        }
        hRow.input.disabled = false;
        wRow.input.disabled = false;
        hRow.input.value = String(Math.round(lay.heightPx));
        wRow.input.value = String(Math.round(lay.widthPx));
      } finally {
        syncingPicSize = false;
      }
    };

    /**
     * 无「锁定横纵比」开关前，默认按当前框比例缩放：以提交侧（宽/高）为主，另一维由比例推出；
     * 若推出值低于最小幅面，则夹紧后回算主维，避免拉伸变形。
     */
    const applyPicSize = (drive: "w" | "h"): void => {
      if (syncingPicSize) {
        return;
      }
      const fs = getFs();
      if (fs === undefined || fs.setSelectedFloatingPictureSizePx === undefined) {
        return;
      }
      const lay = fs.getSelectedFloatingPictureLayoutPx?.() ?? null;
      if (lay === null) {
        return;
      }
      const ow = lay.widthPx;
      const oh = lay.heightPx;
      if (!Number.isFinite(ow) || !Number.isFinite(oh) || ow <= 1e-6 || oh <= 1e-6) {
        return;
      }
      const ar = ow / oh;

      if (drive === "w") {
        const wIn = Number(wRow.input.value);
        if (!Number.isFinite(wIn)) {
          return;
        }
        let w = Math.max(MIN_PIC_PX, Math.round(wIn));
        let h = Math.round(w / ar);
        if (h < MIN_PIC_PX) {
          h = MIN_PIC_PX;
          w = Math.max(MIN_PIC_PX, Math.round(h * ar));
        }
        fs.setSelectedFloatingPictureSizePx(w, h);
        return;
      }

      const hIn = Number(hRow.input.value);
      if (!Number.isFinite(hIn)) {
        return;
      }
      let h = Math.max(MIN_PIC_PX, Math.round(hIn));
      let w = Math.round(h * ar);
      if (w < MIN_PIC_PX) {
        w = MIN_PIC_PX;
        h = Math.max(MIN_PIC_PX, Math.round(w / ar));
      }
      fs.setSelectedFloatingPictureSizePx(w, h);
    };

    hRow.input.addEventListener("change", () => {
      applyPicSize("h");
    });
    wRow.input.addEventListener("change", () => {
      applyPicSize("w");
    });
    hRow.input.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter") {
        ev.preventDefault();
        applyPicSize("h");
      }
    });
    wRow.input.addEventListener("keydown", (ev) => {
      if (ev.key === "Enter") {
        ev.preventDefault();
        applyPicSize("w");
      }
    });

    const fs0 = getFs();
    const unsub = fs0?.subscribeSelectedFloatingPictureLayout?.(() => {
      refreshPicSizeInputs();
    });
    if (unsub !== undefined) {
      cleanups.push(unsub);
    }

    inner.appendChild(sizeRoot);
  }

  return () => {
    for (const c of cleanups) {
      c();
    }
  };
}
