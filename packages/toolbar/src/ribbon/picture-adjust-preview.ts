/**
 * 与 flexsheet `buildFloatingPictureCssFilter` 及 `RECOLOR_EXTRA_CSS` 保持键与顺序一致，供 Ribbon 缩略预览。
 * （toolbar 不依赖 flexsheet 包，故滤镜字符串在此重复维护。）
 */

import type { FloatingPictureAdjustmentsState } from "./ribbon-types.js";

function clampN(n: number, lo: number, hi: number): number {
  return Math.min(hi, Math.max(lo, n));
}

/** 与 flexsheet `floating-picture-layer.ts` 中 `RECOLOR_EXTRA_CSS` 同步 */
const RECOLOR_EXTRA_CSS: Record<string, string> = {
  none: "",
  r_gray: "grayscale(1)",
  r_sepia: "sepia(0.92) saturate(0.22) brightness(1.04)",
  r_washout: "brightness(1.14) contrast(0.72) saturate(0.88)",
  r_bwsoft: "grayscale(1) contrast(1.32) brightness(1.06)",
  r_bwmid: "grayscale(1) contrast(2.15) brightness(0.96)",
  r_bwhard: "grayscale(1) contrast(3.4) brightness(0.86)",
  r_dkgray: "sepia(0.12) saturate(0.12) brightness(0.78) contrast(1.08)",
  r_dkblue: "sepia(1) hue-rotate(185deg) saturate(1.85) brightness(0.66) contrast(1.12)",
  r_dkorange: "sepia(1) hue-rotate(-18deg) saturate(1.75) brightness(0.7) contrast(1.08)",
  r_dksilver: "sepia(0.18) saturate(0.28) brightness(0.8) contrast(1.05)",
  r_dkgold: "sepia(1) hue-rotate(12deg) saturate(1.9) brightness(0.74) contrast(1.1)",
  r_dklblue: "sepia(1) hue-rotate(135deg) saturate(1.55) brightness(0.72) contrast(1.08)",
  r_dkgreen: "sepia(1) hue-rotate(58deg) saturate(1.75) brightness(0.7) contrast(1.08)",
  r_plgray: "sepia(0.08) saturate(0.14) brightness(1.06) contrast(0.98)",
  r_plblue: "sepia(0.38) hue-rotate(178deg) saturate(0.85) brightness(1.1) contrast(0.98)",
  r_plorange: "sepia(0.48) hue-rotate(-14deg) saturate(0.78) brightness(1.1) contrast(0.98)",
  r_plsilver: "sepia(0.07) saturate(0.1) brightness(1.12) contrast(0.97)",
  r_plgold: "sepia(0.52) hue-rotate(8deg) saturate(0.72) brightness(1.1) contrast(0.98)",
  r_pllblue: "sepia(0.28) hue-rotate(158deg) saturate(0.58) brightness(1.12) contrast(0.97)",
  r_plgreen: "sepia(0.36) hue-rotate(52deg) saturate(0.64) brightness(1.1) contrast(0.98)",
};

export function buildFloatingPictureCssFilterPreview(a: FloatingPictureAdjustmentsState): string {
  const b = clampN(1 + a.brightnessPct / 100, 0.05, 3);
  const c0 = clampN(1 + a.contrastPct / 100, 0.05, 3);
  const sat = clampN(a.saturationPct / 100, 0, 4);
  const op = clampN(1 - a.transparencyPct / 100, 0, 1);
  let blurPx = 0;
  let sharpenBoost = 1;
  if (a.sharpnessPct < 0) {
    blurPx = clampN((-a.sharpnessPct / 100) * 3, 0, 3.5);
  } else if (a.sharpnessPct > 0) {
    sharpenBoost = 1 + (a.sharpnessPct / 100) * 0.55;
  }
  const c = c0 * sharpenBoost;
  const parts: string[] = [];
  if (blurPx > 0.05) {
    parts.push(`blur(${blurPx}px)`);
  }
  parts.push(`brightness(${b})`);
  parts.push(`contrast(${c})`);
  parts.push(`saturate(${sat})`);
  const tk = a.colorTemperatureK;
  if (Number.isFinite(tk) && tk !== 6500) {
    const deg = clampN(((tk - 6500) / 4500) * 28, -32, 32);
    if (Math.abs(deg) > 0.3) {
      parts.push(`hue-rotate(${deg}deg)`);
    }
  }
  const recolor = RECOLOR_EXTRA_CSS[a.recolorPreset] ?? "";
  if (recolor !== "") {
    parts.push(recolor);
  }
  parts.push(`opacity(${op})`);
  return parts.join(" ");
}
