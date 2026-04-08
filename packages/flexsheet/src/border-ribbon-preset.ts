import {
  normalizeSelectionRange,
  type CellBorderKind,
  type CellBorderSide,
  type CellStyle,
  type SelectionRange,
  type Worksheet,
} from "@flexsheet/core";

const DEFAULT_BORDER_COLOR = "FF000000";

function side(kind: CellBorderKind, colorArgb: string = DEFAULT_BORDER_COLOR): CellBorderSide {
  return { kind, colorArgb };
}

/** 深拷贝样式中的边框字段，供撤销快照使用。 */
export function cloneCellStyle(s: CellStyle | null): CellStyle | null {
  if (s === null) {
    return null;
  }
  const next: CellStyle = { ...s };
  if (s.borderTop !== undefined) {
    next.borderTop = { ...s.borderTop };
  }
  if (s.borderLeft !== undefined) {
    next.borderLeft = { ...s.borderLeft };
  }
  if (s.borderBottom !== undefined) {
    next.borderBottom = { ...s.borderBottom };
  }
  if (s.borderRight !== undefined) {
    next.borderRight = { ...s.borderRight };
  }
  return next;
}

function clearBorderFields(s: CellStyle | null): CellStyle | null {
  const next = cloneCellStyle(s);
  if (next === null) {
    return null;
  }
  delete next.borderTop;
  delete next.borderLeft;
  delete next.borderBottom;
  delete next.borderRight;
  return Object.keys(next).length === 0 ? null : next;
}

/**
 * 根据 Ribbon 边框命令 id 计算选区内单个非覆盖格（主格）应用后的样式。
 *
 * **几何语义**：`startRow/endRow/startCol/endCol` 为规范化选区矩形。上/下/左/右、底边变体等
 * 只在该矩形**对应外边界**上画线（主格合并矩形贴顶/贴底/贴左/贴右于选区边时才设置该侧）；
 * 「所有框线」为选区内的**完整网格**（矩形内每格四边细线，与 Excel 一致）。
 */
export function computeBorderStyleForRibbonCommand(
  sheet: Worksheet,
  range: SelectionRange,
  row: number,
  col: number,
  commandId: string,
  before: CellStyle | null,
): CellStyle | null {
  const n = normalizeSelectionRange(range);
  const info = sheet.getMergedRectInfo(row, col);
  const ar = info.anchorRow;
  const ac = info.anchorCol;
  const endR = ar + info.rowSpan - 1;
  const endC = ac + info.colSpan - 1;

  const base = cloneCellStyle(before);

  switch (commandId) {
    case "home.font.border":
    case "home.font.border.all": {
      const out = base ?? {};
      const t = side("thin");
      out.borderTop = t;
      out.borderLeft = t;
      out.borderBottom = t;
      out.borderRight = t;
      return out;
    }
    case "home.font.border.none":
      return clearBorderFields(base);
    case "home.font.border.bottom": {
      if (endR !== n.endRow) {
        return base;
      }
      const out = base ?? {};
      out.borderBottom = side("thin");
      return out;
    }
    case "home.font.border.top": {
      if (ar !== n.startRow) {
        return base;
      }
      const out = base ?? {};
      out.borderTop = side("thin");
      return out;
    }
    case "home.font.border.left": {
      if (ac !== n.startCol) {
        return base;
      }
      const out = base ?? {};
      out.borderLeft = side("thin");
      return out;
    }
    case "home.font.border.right": {
      if (endC !== n.endCol) {
        return base;
      }
      const out = base ?? {};
      out.borderRight = side("thin");
      return out;
    }
    case "home.font.border.outside": {
      const out = base ?? {};
      const t = side("thin");
      if (ar === n.startRow) {
        out.borderTop = t;
      }
      if (endR === n.endRow) {
        out.borderBottom = t;
      }
      if (ac === n.startCol) {
        out.borderLeft = t;
      }
      if (endC === n.endCol) {
        out.borderRight = t;
      }
      return Object.keys(out).length === 0 ? null : out;
    }
    case "home.font.border.thickBox": {
      const out = base ?? {};
      const t = side("thick");
      if (ar === n.startRow) {
        out.borderTop = t;
      }
      if (endR === n.endRow) {
        out.borderBottom = t;
      }
      if (ac === n.startCol) {
        out.borderLeft = t;
      }
      if (endC === n.endCol) {
        out.borderRight = t;
      }
      return Object.keys(out).length === 0 ? null : out;
    }
    case "home.font.border.doubleBottom": {
      if (endR !== n.endRow) {
        return base;
      }
      const out = base ?? {};
      out.borderBottom = side("double");
      return out;
    }
    case "home.font.border.thickBottom": {
      if (endR !== n.endRow) {
        return base;
      }
      const out = base ?? {};
      out.borderBottom = side("thick");
      return out;
    }
    case "home.font.border.topBottom": {
      const out = base ?? {};
      const t = side("thin");
      let touched = false;
      if (ar === n.startRow) {
        out.borderTop = t;
        touched = true;
      }
      if (endR === n.endRow) {
        out.borderBottom = t;
        touched = true;
      }
      if (!touched) {
        return base;
      }
      return out;
    }
    case "home.font.border.topThickBottom": {
      const out = base ?? {};
      let touched = false;
      if (ar === n.startRow) {
        out.borderTop = side("thin");
        touched = true;
      }
      if (endR === n.endRow) {
        out.borderBottom = side("thick");
        touched = true;
      }
      if (!touched) {
        return base;
      }
      return out;
    }
    case "home.font.border.topDoubleBottom": {
      const out = base ?? {};
      let touched = false;
      if (ar === n.startRow) {
        out.borderTop = side("thin");
        touched = true;
      }
      if (endR === n.endRow) {
        out.borderBottom = side("double");
        touched = true;
      }
      if (!touched) {
        return base;
      }
      return out;
    }
    default:
      return base;
  }
}
