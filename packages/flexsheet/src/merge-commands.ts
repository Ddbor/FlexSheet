import {
  Cell,
  normalizeSelectionRange,
  type CellScalar,
  type CellStyle,
  type ICommand,
  type SelectionRange,
  type Worksheet,
} from "@flexsheet/core";

type CellSnap = {
  readonly value: CellScalar;
  readonly formula: string | null;
  readonly style: CellStyle | null;
};

type MergeSnap = {
  readonly masterRow: number;
  readonly masterCol: number;
  readonly rowSpan: number;
  readonly colSpan: number;
};

const LABELS: Record<"mergeCells" | "mergeAcross" | "mergeCenter" | "unmerge", string> = {
  mergeCells: "合并单元格",
  mergeAcross: "跨越合并",
  mergeCenter: "合并后居中",
  unmerge: "取消单元格合并",
};

function parseCellKey(key: string): [number, number] {
  const i = key.indexOf(",");
  return [Number(key.slice(0, i)), Number(key.slice(i + 1))];
}

function snapshotCells(sheet: Worksheet, range: SelectionRange): Map<string, CellSnap> {
  const n = normalizeSelectionRange(range);
  const map = new Map<string, CellSnap>();
  for (let r = n.startRow; r <= n.endRow; r++) {
    for (let c = n.startCol; c <= n.endCol; c++) {
      const cell = sheet.getCell(r, c);
      map.set(Cell.key(r, c), {
        value: cell.value,
        formula: cell.formula,
        style: cell.style === null ? null : { ...cell.style },
      });
    }
  }
  return map;
}

function restoreCells(sheet: Worksheet, cells: ReadonlyMap<string, CellSnap>): void {
  sheet.batch(() => {
    for (const [k, snap] of cells) {
      const [r, c] = parseCellKey(k);
      if (snap.formula !== null && snap.formula.length > 0) {
        sheet.setCellFormula(r, c, snap.formula);
      } else {
        sheet.setCellLiteral(r, c, snap.value);
      }
      sheet.setCellStyle(r, c, snap.style);
    }
  });
}

function cloneMergeSnap(sheet: Worksheet): MergeSnap[] {
  return sheet.getMergeRegionsSnapshot().map((r) => ({
    masterRow: r.masterRow,
    masterCol: r.masterCol,
    rowSpan: r.rowSpan,
    colSpan: r.colSpan,
  }));
}

/**
 * Ribbon 合并/取消合并（可撤销/重做）：快照合并映射与选区内单元格内容。
 */
export class SelectionMergeCommand implements ICommand {
  readonly id = "sheet.selectionMerge";
  readonly label: string;
  private readonly backwardRegions: MergeSnap[];
  private readonly backwardCells: Map<string, CellSnap>;
  private forwardRegions: MergeSnap[] | null = null;
  private forwardCells: Map<string, CellSnap> | null = null;

  constructor(
    private readonly sheet: Worksheet,
    private readonly range: SelectionRange,
    private readonly kind: "mergeCells" | "mergeAcross" | "mergeCenter" | "unmerge",
  ) {
    this.label = LABELS[kind];
    this.backwardRegions = cloneMergeSnap(sheet);
    this.backwardCells = snapshotCells(sheet, range);
  }

  execute(): void {
    if (this.forwardRegions !== null && this.forwardCells !== null) {
      this.sheet.restoreMergeRegionsFromSnapshot(this.forwardRegions);
      restoreCells(this.sheet, this.forwardCells);
      return;
    }
    if (this.kind === "unmerge") {
      this.sheet.applyUnmergeForSelection(this.range);
    } else {
      this.sheet.applyMergeForSelection(this.range, this.kind);
    }
    this.forwardRegions = cloneMergeSnap(this.sheet);
    this.forwardCells = snapshotCells(this.sheet, this.range);
  }

  undo(): void {
    this.sheet.restoreMergeRegionsFromSnapshot(this.backwardRegions);
    restoreCells(this.sheet, this.backwardCells);
  }
}
