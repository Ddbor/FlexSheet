import type { CellBorderKind, CellBorderSide } from "./cell.js";

export interface ResolvedCellBorderStroke {
  readonly lineWidth: number;
  /** 双线间距（仅 double 时 >0） */
  readonly gap: number;
  readonly lineDash: readonly number[] | null;
  readonly double: boolean;
  readonly lineCap: CanvasLineCap;
}

function legacyFromKind(kind: CellBorderKind, z: number): ResolvedCellBorderStroke {
  switch (kind) {
    case "hairline":
      return {
        lineWidth: Math.max(0.5, 0.5 * z),
        gap: 0,
        lineDash: null,
        double: false,
        lineCap: "butt",
      };
    case "thin":
      return {
        lineWidth: Math.max(1, z),
        gap: 0,
        lineDash: null,
        double: false,
        lineCap: "butt",
      };
    case "medium":
      return {
        lineWidth: Math.max(2, 2 * z),
        gap: 0,
        lineDash: null,
        double: false,
        lineCap: "butt",
      };
    case "thick":
      return {
        lineWidth: Math.max(3, 3 * z),
        gap: 0,
        lineDash: null,
        double: false,
        lineCap: "butt",
      };
    case "double":
      return {
        lineWidth: Math.max(1, z),
        gap: Math.max(2, 2 * z),
        lineDash: null,
        double: true,
        lineCap: "butt",
      };
    default:
      return {
        lineWidth: Math.max(1, z),
        gap: 0,
        lineDash: null,
        double: false,
        lineCap: "butt",
      };
  }
}

/**
 * 将单边边框解析为 Canvas 描边参数（线宽、虚线、双线）。
 * 线型网格左列为细、右列为粗：细线型用较小 lineWidth，粗线型用明显更大线宽，避免观感一致。
 */
export function resolveCellBorderStroke(side: CellBorderSide, viewZoom: number): ResolvedCellBorderStroke {
  const z = Math.max(0.25, viewZoom);
  if (side.linePattern === undefined) {
    return legacyFromKind(side.kind, z);
  }
  const p = side.linePattern;
  switch (p) {
    case "hairlineDots":
      return {
        lineWidth: Math.max(0.42, 0.42 * z),
        gap: 0,
        lineDash: [1 * z, 5 * z],
        double: false,
        lineCap: "round",
      };
    case "shortDash":
      return {
        lineWidth: Math.max(0.78, 0.78 * z),
        gap: 0,
        lineDash: [6 * z, 4 * z],
        double: false,
        lineCap: "butt",
      };
    case "dashDot":
      return {
        lineWidth: Math.max(0.72, 0.72 * z),
        gap: 0,
        lineDash: [10 * z, 3 * z, 2 * z, 3 * z],
        double: false,
        lineCap: "butt",
      };
    case "dashDotDot":
      return {
        lineWidth: Math.max(0.76, 0.76 * z),
        gap: 0,
        lineDash: [10 * z, 3 * z, 2 * z, 3 * z, 2 * z, 3 * z],
        double: false,
        lineCap: "butt",
      };
    case "mediumDash":
      return {
        lineWidth: Math.max(1.02, 1.02 * z),
        gap: 0,
        lineDash: [8 * z, 4 * z],
        double: false,
        lineCap: "butt",
      };
    case "mediumDashDotDot":
      return {
        lineWidth: Math.max(2.55, 2.55 * z),
        gap: 0,
        lineDash: [14 * z, 3 * z, 2 * z, 3 * z, 2 * z, 3 * z],
        double: false,
        lineCap: "butt",
      };
    case "slantedDash":
      return {
        lineWidth: Math.max(1.05, 1.05 * z),
        gap: 0,
        lineDash: [4 * z, 3 * z],
        double: false,
        lineCap: "round",
      };
    case "thickDash":
      return {
        lineWidth: Math.max(3.1, 3.1 * z),
        gap: 0,
        lineDash: [6 * z, 4 * z],
        double: false,
        lineCap: "butt",
      };
    case "thickDashDot":
      return {
        lineWidth: Math.max(2.85, 2.85 * z),
        gap: 0,
        lineDash: [12 * z, 3 * z, 2 * z, 3 * z],
        double: false,
        lineCap: "butt",
      };
    case "thinSolid":
      return {
        lineWidth: Math.max(1.02, 1.02 * z),
        gap: 0,
        lineDash: null,
        double: false,
        lineCap: "butt",
      };
    case "mediumSolid":
      return {
        lineWidth: Math.max(2.2, 2.2 * z),
        gap: 0,
        lineDash: null,
        double: false,
        lineCap: "butt",
      };
    case "thickSolid":
      return {
        lineWidth: Math.max(3.55, 3.55 * z),
        gap: 0,
        lineDash: null,
        double: false,
        lineCap: "butt",
      };
    case "doubleLine":
      return {
        lineWidth: Math.max(0.95, 0.95 * z),
        gap: Math.max(2.5, 2.5 * z),
        lineDash: null,
        double: true,
        lineCap: "butt",
      };
    default: {
      const _x: never = p;
      void _x;
      return legacyFromKind(side.kind, z);
    }
  }
}
