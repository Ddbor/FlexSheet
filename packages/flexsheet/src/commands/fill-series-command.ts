import type { CellScalar, ICommand, SelectionRange, Worksheet } from "@flexsheet/core";
import { excelSerialToUtcDate, normalizeSelectionRange } from "@flexsheet/core";
import { recalcWorksheet } from "@flexsheet/formula";

interface CellValueSnapshot {
  readonly row: number;
  readonly col: number;
  readonly beforeFormula: string | null;
  readonly beforeValue: CellScalar;
  readonly afterValue: CellScalar;
}

export interface FillSeriesOptions {
  readonly seriesIn: "row" | "col";
  readonly type: "linear" | "growth" | "date" | "autofill";
  readonly dateUnit: "day" | "weekday" | "month" | "year";
  readonly step: number;
  readonly stop: number | null;
  readonly trend: boolean;
}

function scalarToNumber(v: CellScalar): number | null {
  if (typeof v === "number" && Number.isFinite(v)) {
    return v;
  }
  if (typeof v === "string") {
    const n = Number(v.trim());
    return Number.isFinite(n) ? n : null;
  }
  return null;
}

function scalarToDateSerial(v: CellScalar): number | null {
  if (typeof v === "number" && Number.isFinite(v)) {
    return v;
  }
  if (typeof v === "string") {
    const t = v.trim();
    const n = Number(t);
    if (Number.isFinite(n)) {
      return n;
    }
    const ms = Date.parse(t);
    if (Number.isFinite(ms)) {
      return ms / 86400000 + 25569;
    }
  }
  return null;
}

function addDateSerial(serial: number, unit: FillSeriesOptions["dateUnit"], step: number): number {
  if (unit === "day") {
    return serial + step;
  }
  if (unit === "weekday") {
    return addWeekdaySerial(serial, step);
  }
  const d = excelSerialToUtcDate(serial);
  if (unit === "month") {
    d.setUTCMonth(d.getUTCMonth() + Math.trunc(step));
  } else {
    d.setUTCFullYear(d.getUTCFullYear() + Math.trunc(step));
  }
  return d.getTime() / 86400000 + 25569;
}

function addWeekdaySerial(serial: number, step: number): number {
  const signedDays = Math.trunc(step);
  if (signedDays === 0) {
    return serial;
  }
  const dir = signedDays > 0 ? 1 : -1;
  let remaining = Math.abs(signedDays);
  let d = excelSerialToUtcDate(serial);
  while (remaining > 0) {
    d = new Date(d.getTime() + dir * 86400000);
    const wd = d.getUTCDay();
    if (wd !== 0 && wd !== 6) {
      remaining--;
    }
  }
  return d.getTime() / 86400000 + 25569;
}

function shouldStop(prev: number, next: number, stop: number | null): boolean {
  if (stop === null || !Number.isFinite(stop)) {
    return false;
  }
  if (next === prev) {
    return false;
  }
  if (next > prev) {
    return next > stop;
  }
  return next < stop;
}

function deriveLinearStep(seedNums: readonly (number | null)[], fallback: number): number {
  const known = seedNums.filter((n): n is number => n !== null);
  if (known.length >= 2) {
    return known[1] - known[0];
  }
  return fallback;
}

function deriveGrowthFactor(seedNums: readonly (number | null)[], fallback: number): number {
  const known = seedNums.filter((n): n is number => n !== null);
  if (known.length >= 2 && known[0] !== 0) {
    return known[1] / known[0];
  }
  return fallback;
}

/** 选区「填充 -> 系列」：按行/列方向对当前选区写入序列值（可撤销）。 */
export class FillSeriesCommand implements ICommand {
  readonly id = "sheet.fillSeries";
  readonly label = "系列填充";
  private readonly snapshots: CellValueSnapshot[] = [];

  constructor(
    private readonly sheet: Worksheet,
    range: SelectionRange,
    private readonly options: FillSeriesOptions,
  ) {
    const n = normalizeSelectionRange(range);
    if (options.seriesIn === "row") {
      for (let r = n.startRow; r <= n.endRow; r++) {
        const cols: number[] = [];
        for (let c = n.startCol; c <= n.endCol; c++) {
          cols.push(c);
        }
        this.prepareLine(r, cols, "row");
      }
    } else {
      for (let c = n.startCol; c <= n.endCol; c++) {
        const rows: number[] = [];
        for (let r = n.startRow; r <= n.endRow; r++) {
          rows.push(r);
        }
        this.prepareLine(c, rows, "col");
      }
    }
  }

  private prepareLine(
    fixed: number,
    varying: readonly number[],
    mode: "row" | "col",
  ): void {
    if (varying.length <= 1) {
      return;
    }
    const coords = varying.map((v) =>
      mode === "row" ? { row: fixed, col: v } : { row: v, col: fixed },
    );
    const seeds = coords.map((p) => this.sheet.getCell(p.row, p.col).value);
    const seedNums = seeds.map(scalarToNumber);
    const seedDates = seeds.map(scalarToDateSerial);

    if (this.options.type === "autofill") {
      this.prepareAutofill(coords, seeds);
      return;
    }
    if (this.options.type === "growth") {
      this.prepareGrowth(coords, seedNums);
      return;
    }
    if (this.options.type === "date") {
      this.prepareDate(coords, seedDates);
      return;
    }
    this.prepareLinear(coords, seedNums);
  }

  private prepareAutofill(
    coords: readonly { row: number; col: number }[],
    seeds: readonly CellScalar[],
  ): void {
    const first = seeds[0] ?? null;
    const firstNum = scalarToNumber(first);
    const secondNum = scalarToNumber(seeds[1] ?? null);
    const useLinear = firstNum !== null && secondNum !== null;
    const step = useLinear ? secondNum - firstNum : this.options.step;
    let prev = firstNum ?? 0;
    for (let i = 1; i < coords.length; i++) {
      const p = coords[i]!;
      const cell = this.sheet.getCell(p.row, p.col);
      const next = useLinear ? prev + step : prev + this.options.step;
      const after: CellScalar = useLinear ? next : first;
      this.snapshots.push({
        row: p.row,
        col: p.col,
        beforeFormula: cell.formula,
        beforeValue: cell.value,
        afterValue: after,
      });
      prev = next;
    }
  }

  private prepareLinear(
    coords: readonly { row: number; col: number }[],
    seedNums: readonly (number | null)[],
  ): void {
    const base = seedNums[0];
    if (base === null) {
      return;
    }
    const step = this.options.trend
      ? deriveLinearStep(seedNums, this.options.step)
      : this.options.step;
    let prev = base;
    for (let i = 1; i < coords.length; i++) {
      const next = prev + step;
      if (shouldStop(prev, next, this.options.stop)) {
        break;
      }
      const p = coords[i]!;
      const cell = this.sheet.getCell(p.row, p.col);
      this.snapshots.push({
        row: p.row,
        col: p.col,
        beforeFormula: cell.formula,
        beforeValue: cell.value,
        afterValue: next,
      });
      prev = next;
    }
  }

  private prepareGrowth(
    coords: readonly { row: number; col: number }[],
    seedNums: readonly (number | null)[],
  ): void {
    const base = seedNums[0];
    if (base === null) {
      return;
    }
    const factor = this.options.trend
      ? deriveGrowthFactor(seedNums, this.options.step)
      : this.options.step;
    let prev = base;
    for (let i = 1; i < coords.length; i++) {
      const next = prev * factor;
      if (shouldStop(prev, next, this.options.stop)) {
        break;
      }
      const p = coords[i]!;
      const cell = this.sheet.getCell(p.row, p.col);
      this.snapshots.push({
        row: p.row,
        col: p.col,
        beforeFormula: cell.formula,
        beforeValue: cell.value,
        afterValue: next,
      });
      prev = next;
    }
  }

  private prepareDate(
    coords: readonly { row: number; col: number }[],
    seedDates: readonly (number | null)[],
  ): void {
    const base = seedDates[0];
    if (base === null) {
      return;
    }
    const step = this.options.trend
      ? deriveLinearStep(seedDates, this.options.step)
      : this.options.step;
    let prev = base;
    for (let i = 1; i < coords.length; i++) {
      const next = addDateSerial(prev, this.options.dateUnit, step);
      if (shouldStop(prev, next, this.options.stop)) {
        break;
      }
      const p = coords[i]!;
      const cell = this.sheet.getCell(p.row, p.col);
      this.snapshots.push({
        row: p.row,
        col: p.col,
        beforeFormula: cell.formula,
        beforeValue: cell.value,
        afterValue: next,
      });
      prev = next;
    }
  }

  execute(): void {
    if (this.snapshots.length === 0) {
      return;
    }
    this.sheet.batch(() => {
      for (const s of this.snapshots) {
        this.sheet.setCellLiteral(s.row, s.col, s.afterValue);
      }
    });
    recalcWorksheet(this.sheet);
  }

  undo(): void {
    if (this.snapshots.length === 0) {
      return;
    }
    this.sheet.batch(() => {
      for (const s of this.snapshots) {
        if (s.beforeFormula !== null && s.beforeFormula.length > 0) {
          this.sheet.setCellFormula(s.row, s.col, s.beforeFormula);
        } else {
          this.sheet.setCellLiteral(s.row, s.col, s.beforeValue);
        }
      }
    });
    recalcWorksheet(this.sheet);
  }
}
