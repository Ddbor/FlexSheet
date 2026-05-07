import type {
  FloatingPictureGradientStop,
  FloatingPictureGradientType,
  FloatingPictureFrameFill,
} from "./floating-picture-layer.js";

/** 线性渐变 8 方向：用户角度 0°=左→右，90°=上→下，顺时针。索引与方向面板从左到右、从上到下对应。 */
export const LINEAR_DIRECTION_USER_ANGLES: readonly number[] = [315, 0, 45, 90, 135, 180, 225, 270];

export interface FormatPictureGradientPresetDef {
  readonly id: number;
  readonly gradientType: FloatingPictureGradientType;
  readonly gradientAngleDeg: number;
  readonly stops: readonly FloatingPictureGradientStop[];
}

function stops2(a: string, b: string, mid = 50): readonly FloatingPictureGradientStop[] {
  return [
    { positionPct: 0, color: a, transparencyPct: 0, brightnessPct: 0 },
    { positionPct: mid, color: b, transparencyPct: 0, brightnessPct: 0 },
    { positionPct: 100, color: a, transparencyPct: 0, brightnessPct: 0 },
  ];
}

function stops2simple(a: string, b: string): readonly FloatingPictureGradientStop[] {
  return [
    { positionPct: 0, color: a, transparencyPct: 0, brightnessPct: 0 },
    { positionPct: 100, color: b, transparencyPct: 0, brightnessPct: 0 },
  ];
}

/** 6×5 预设渐变（与 Office 色列大致对应：蓝、橙、灰、金、浅蓝、绿）。 */
export const FORMAT_PICTURE_GRADIENT_PRESETS: readonly FormatPictureGradientPresetDef[] = [
  // 蓝系
  {
    id: 0,
    gradientType: "linear",
    gradientAngleDeg: 90,
    stops: stops2simple("#5b9bd5", "#ffffff"),
  },
  { id: 1, gradientType: "linear", gradientAngleDeg: 0, stops: stops2simple("#5b9bd5", "#ffffff") },
  {
    id: 2,
    gradientType: "radial",
    gradientAngleDeg: 90,
    stops: stops2simple("#2e75b6", "#bdd7ee"),
  },
  { id: 3, gradientType: "linear", gradientAngleDeg: 45, stops: stops2("#4472c4", "#ffffff", 40) },
  {
    id: 4,
    gradientType: "radial",
    gradientAngleDeg: 90,
    stops: stops2simple("#1f4e79", "#9dc3e6"),
  },
  // 橙系
  {
    id: 5,
    gradientType: "linear",
    gradientAngleDeg: 90,
    stops: stops2simple("#ed7d31", "#fce4d6"),
  },
  { id: 6, gradientType: "linear", gradientAngleDeg: 0, stops: stops2simple("#c65911", "#f8cbad") },
  {
    id: 7,
    gradientType: "radial",
    gradientAngleDeg: 90,
    stops: stops2simple("#843c0c", "#f4b183"),
  },
  { id: 8, gradientType: "linear", gradientAngleDeg: 135, stops: stops2("#ed7d31", "#fff2cc", 35) },
  {
    id: 9,
    gradientType: "radial",
    gradientAngleDeg: 90,
    stops: stops2simple("#bf8f00", "#fff2cc"),
  },
  // 灰系
  {
    id: 10,
    gradientType: "linear",
    gradientAngleDeg: 90,
    stops: stops2simple("#7f7f7f", "#f2f2f2"),
  },
  {
    id: 11,
    gradientType: "linear",
    gradientAngleDeg: 0,
    stops: stops2simple("#595959", "#d9d9d9"),
  },
  {
    id: 12,
    gradientType: "radial",
    gradientAngleDeg: 90,
    stops: stops2simple("#3f3f3f", "#bfbfbf"),
  },
  { id: 13, gradientType: "linear", gradientAngleDeg: 45, stops: stops2("#a6a6a6", "#ffffff", 50) },
  {
    id: 14,
    gradientType: "radial",
    gradientAngleDeg: 90,
    stops: stops2simple("#404040", "#d0d0d0"),
  },
  // 金 / 黄系
  {
    id: 15,
    gradientType: "linear",
    gradientAngleDeg: 90,
    stops: stops2simple("#ffc000", "#fff2cc"),
  },
  {
    id: 16,
    gradientType: "linear",
    gradientAngleDeg: 0,
    stops: stops2simple("#bf9000", "#ffe699"),
  },
  {
    id: 17,
    gradientType: "radial",
    gradientAngleDeg: 90,
    stops: stops2simple("#806000", "#ffd966"),
  },
  {
    id: 18,
    gradientType: "linear",
    gradientAngleDeg: 315,
    stops: stops2("#ffc000", "#ffffff", 45),
  },
  {
    id: 19,
    gradientType: "radial",
    gradientAngleDeg: 90,
    stops: stops2simple("#997300", "#fff9e6"),
  },
  // 浅蓝系
  {
    id: 20,
    gradientType: "linear",
    gradientAngleDeg: 90,
    stops: stops2simple("#9dc3e6", "#ffffff"),
  },
  {
    id: 21,
    gradientType: "linear",
    gradientAngleDeg: 0,
    stops: stops2simple("#8faadc", "#e7e6e6"),
  },
  {
    id: 22,
    gradientType: "radial",
    gradientAngleDeg: 90,
    stops: stops2simple("#2e75b6", "#deebf7"),
  },
  { id: 23, gradientType: "linear", gradientAngleDeg: 90, stops: stops2("#bdd7ee", "#ffffff", 60) },
  {
    id: 24,
    gradientType: "radial",
    gradientAngleDeg: 90,
    stops: stops2simple("#1f4e79", "#d6dce4"),
  },
  // 绿系
  {
    id: 25,
    gradientType: "linear",
    gradientAngleDeg: 90,
    stops: stops2simple("#70ad47", "#e2efda"),
  },
  {
    id: 26,
    gradientType: "linear",
    gradientAngleDeg: 0,
    stops: stops2simple("#548235", "#c6e0b4"),
  },
  {
    id: 27,
    gradientType: "radial",
    gradientAngleDeg: 90,
    stops: stops2simple("#375623", "#a9d08e"),
  },
  { id: 28, gradientType: "linear", gradientAngleDeg: 45, stops: stops2("#70ad47", "#ffffff", 40) },
  {
    id: 29,
    gradientType: "radial",
    gradientAngleDeg: 90,
    stops: stops2simple("#264d3a", "#c5e0b4"),
  },
];

export function presetToFrameFillPatch(
  p: FormatPictureGradientPresetDef,
): Partial<FloatingPictureFrameFill> {
  const idx = LINEAR_DIRECTION_USER_ANGLES.indexOf(p.gradientAngleDeg);
  const patch: Partial<FloatingPictureFrameFill> = {
    kind: "gradient",
    gradientType: p.gradientType,
    gradientAngleDeg: p.gradientAngleDeg,
    linearDirectionIndex: idx >= 0 ? idx : undefined,
    gradientStops: p.stops.map((s) => ({ ...s })),
    gradientPresetId: p.id,
  };
  if (p.gradientType === "radial") {
    patch.radialFillLtrb = { l: 0, t: 0, r: 100000, b: 100000 };
    patch.radialTileLtrb = { l: -100000, t: -100000, r: 0, b: 0 };
  }
  return patch;
}

/** 用于下拉缩略图：生成 background-image CSS 片段（不含分号）。 */
export function presetThumbnailBackground(p: FormatPictureGradientPresetDef): string {
  const cssAng = userGradientAngleToCssAngle(p.gradientAngleDeg);
  const parts = [...p.stops]
    .sort((a, b) => a.positionPct - b.positionPct)
    .map((s) => {
      const rgb = hexToRgb(s.color);
      const f = Math.min(2, Math.max(0, 1 + s.brightnessPct / 100));
      const a = 1 - Math.min(100, Math.max(0, s.transparencyPct)) / 100;
      const r = Math.round(clamp255(rgb.r * f));
      const g = Math.round(clamp255(rgb.g * f));
      const b = Math.round(clamp255(rgb.b * f));
      return `rgba(${r},${g},${b},${a}) ${s.positionPct}%`;
    });
  if (p.gradientType === "linear") {
    return `linear-gradient(${cssAng}deg, ${parts.join(", ")})`;
  }
  return `radial-gradient(circle at center, ${parts.join(", ")})`;
}

function userGradientAngleToCssAngle(userDeg: number): number {
  const u = ((userDeg % 360) + 360) % 360;
  return (90 + u) % 360;
}

function hexToRgb(hex: string): { readonly r: number; readonly g: number; readonly b: number } {
  const h = hex.trim();
  const m = /^#?([0-9a-f]{2})([0-9a-f]{2})([0-9a-f]{2})$/i.exec(h);
  if (m === null) {
    return { r: 0, g: 0, b: 0 };
  }
  return {
    r: Number.parseInt(m[1]!, 16),
    g: Number.parseInt(m[2]!, 16),
    b: Number.parseInt(m[3]!, 16),
  };
}

function clamp255(n: number): number {
  return Math.min(255, Math.max(0, n));
}
