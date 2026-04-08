import type { CellStyle } from "@flexsheet/core";
import { argb8ToStripeCss } from "./ribbon-color-argb.js";
import {
  RIBBON_FONT_FAMILY_DEFAULT_PREVIEW,
  RIBBON_FONT_FAMILY_ITEMS,
} from "./font-family-items.js";

/** 「开始」字体组控件展示状态（与活动单元格 `CellStyle` 对齐）。 */
export interface RibbonHomeFontChromeState {
  readonly fontLabel: string;
  readonly fontPreviewCss: string;
  readonly sizeLabel: string;
  readonly boldPressed: boolean;
  readonly italicPressed: boolean;
  readonly underlinePressed: boolean;
  readonly doubleUnderlinePressed: boolean;
  /** 填充色条 CSS 颜色；`null` 仅表示全透明 ARGB（透明格图案）。未设 `fillArgb` 时为默认黑底 `#000000`。 */
  readonly fillStripeCss: string | null;
  /** 字体色条；未设 `fgArgb` 时为默认白字 `#ffffff`（与默认黑底对应）。 */
  readonly fontStripeCss: string;
}

function normalizeFontStack(s: string): string {
  return s.replace(/\s+/g, " ").trim().toLowerCase();
}

function itemPreviewCss(it: (typeof RIBBON_FONT_FAMILY_ITEMS)[number]): string {
  if (it.previewFontFamily !== undefined && it.previewFontFamily !== "") {
    return it.previewFontFamily;
  }
  return it.id === "home.font.family.wingdings"
    ? "Wingdings, fantasy"
    : `"${it.label}", sans-serif`;
}

const FONT_STACK_TO_UI = new Map<string, { label: string; preview: string }>();
for (const it of RIBBON_FONT_FAMILY_ITEMS) {
  const preview = itemPreviewCss(it);
  FONT_STACK_TO_UI.set(normalizeFontStack(preview), { label: it.label, preview });
}

function extractDisplayFamilyName(stack: string): string {
  const q = stack.match(/"([^"]+)"/);
  if (q !== null && q[1] !== undefined) {
    return q[1].trim();
  }
  const first = stack.split(",")[0]?.trim();
  if (first !== undefined && first !== "") {
    return first;
  }
  const t = stack.trim();
  return t.length > 0 ? t.slice(0, 16) : "字体";
}

const FILL_STRIPE_DEFAULT = "#000000";
const FONT_STRIPE_DEFAULT = "#ffffff";

function stripeFields(style: CellStyle | null): {
  fillStripeCss: string | null;
  fontStripeCss: string;
} {
  const fa = style?.fillArgb?.trim();
  let fillStripeCss: string | null;
  if (fa === undefined || fa === "") {
    fillStripeCss = FILL_STRIPE_DEFAULT;
  } else if (!/^[\dA-Fa-f]{8}$/.test(fa)) {
    fillStripeCss = FILL_STRIPE_DEFAULT;
  } else {
    const c = argb8ToStripeCss(fa);
    fillStripeCss = c === null ? null : c;
  }

  const fg = style?.fgArgb?.trim();
  let fontStripeCss: string;
  if (fg === undefined || fg === "" || !/^[\dA-Fa-f]{8}$/.test(fg)) {
    fontStripeCss = FONT_STRIPE_DEFAULT;
  } else {
    fontStripeCss = argb8ToStripeCss(fg) ?? FONT_STRIPE_DEFAULT;
  }

  return { fillStripeCss, fontStripeCss };
}

/** 由活动单元格样式推导 Ribbon 字体下拉与切换按钮状态。 */
export function cellStyleToRibbonHomeFontChrome(
  style: CellStyle | null,
): RibbonHomeFontChromeState {
  const boldPressed = style?.bold === true;
  const italicPressed = style?.italic === true;
  const underlinePressed = style?.underline === "single";
  const doubleUnderlinePressed = style?.underline === "double";
  const sizeLabel = String(style?.fontSizePt ?? 11);
  const stripes = stripeFields(style);

  const rawFam = style?.fontFamily;
  if (rawFam !== undefined && rawFam.trim() !== "") {
    const hit = FONT_STACK_TO_UI.get(normalizeFontStack(rawFam));
    if (hit !== undefined) {
      return {
        fontLabel: hit.label,
        fontPreviewCss: hit.preview,
        sizeLabel,
        boldPressed,
        italicPressed,
        underlinePressed,
        doubleUnderlinePressed,
        ...stripes,
      };
    }
    return {
      fontLabel: extractDisplayFamilyName(rawFam),
      fontPreviewCss: rawFam,
      sizeLabel,
      boldPressed,
      italicPressed,
      underlinePressed,
      doubleUnderlinePressed,
      ...stripes,
    };
  }

  return {
    fontLabel: "微软雅黑",
    fontPreviewCss: RIBBON_FONT_FAMILY_DEFAULT_PREVIEW,
    sizeLabel,
    boldPressed,
    italicPressed,
    underlinePressed,
    doubleUnderlinePressed,
    ...stripes,
  };
}
