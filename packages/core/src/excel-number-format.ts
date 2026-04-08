import type { CellScalar, CellStyle } from "./cell.js";

/** Excel 序列日期转 UTC `Date`（与 Excel 1900 日期系一致；`25569` 为 1970-01-01）。 */
export function excelSerialToUtcDate(serial: number): Date {
  return new Date(Math.round((serial - 25569) * 86400 * 1000));
}

function pad2(n: number): string {
  return n < 10 ? `0${n}` : String(n);
}

function formatDateLikeExcel(value: number, format: string): string {
  const d = excelSerialToUtcDate(value);
  const y = d.getUTCFullYear();
  const m = d.getUTCMonth() + 1;
  const day = d.getUTCDate();
  const h = d.getUTCHours();
  const min = d.getUTCMinutes();
  const s = d.getUTCSeconds();

  if (/[hH]:/.test(format) || format.includes("秒") || format.includes("mm:ss")) {
    const hh = pad2(h);
    const mm = pad2(min);
    const ss = pad2(s);
    if (format.includes("秒") || format.split(":").length >= 3) {
      return `${hh}:${mm}:${ss}`;
    }
    return `${hh}:${mm}`;
  }

  if (format.includes("年") && format.includes("月")) {
    return `${y}年${m}月${day}日`;
  }
  if (/y{2,4}.*m{1,2}.*d{1,2}/i.test(format) || format.includes("短")) {
    return `${y}/${m}/${day}`;
  }
  return `${y}/${m}/${day}`;
}

function countFractionDigitsInPattern(part: string): number {
  const dot = part.lastIndexOf(".");
  if (dot < 0) {
    return 0;
  }
  let i = dot + 1;
  let n = 0;
  while (i < part.length && part[i] === "0") {
    n++;
    i++;
  }
  return n;
}

function stripPercentSuffix(s: string): { body: string; hasPct: boolean } {
  const t = s.trimEnd();
  if (t.endsWith("%")) {
    return { body: t.slice(0, -1).trimEnd(), hasPct: true };
  }
  return { body: t, hasPct: false };
}

/**
 * 在首段正数样式上增加或减少小数位（与 Excel「增加/减少小数位数」行为接近）。
 */
export function adjustDecimalPlacesInFormat(code: string | undefined, delta: number): string {
  const raw = code?.trim() ?? "";
  if (raw === "" || raw === "General" || raw === "@") {
    if (delta > 0) {
      return delta === 1 ? "0.0" : `0.${"0".repeat(delta)}`;
    }
    return "";
  }

  const sections = raw.split(";");
  const first = sections[0] ?? "";
  const { body, hasPct } = stripPercentSuffix(first);
  const pctSuffix = hasPct ? "%" : "";

  const dot = body.lastIndexOf(".");
  if (delta > 0) {
    if (dot < 0) {
      const nextFirst = `${body}.0${pctSuffix}`;
      return [nextFirst, ...sections.slice(1)].join(";");
    }
    const intPart = body.slice(0, dot);
    const decPart = body.slice(dot + 1);
    const nextFirst = `${intPart}.${decPart}0${pctSuffix}`;
    return [nextFirst, ...sections.slice(1)].join(";");
  }

  if (delta < 0) {
    if (dot < 0) {
      return raw;
    }
    const intPart = body.slice(0, dot);
    const decPart = body.slice(dot + 1);
    if (decPart.length <= 1) {
      const nextBody = intPart;
      const nextFirst = `${nextBody}${pctSuffix}`;
      return [nextFirst, ...sections.slice(1)].join(";");
    }
    const nextFirst = `${intPart}.${decPart.slice(0, -1)}${pctSuffix}`;
    return [nextFirst, ...sections.slice(1)].join(";");
  }

  return raw;
}

/** 千位分隔样式：在保留现有小数位意图的前提下套用 `#,##0` 系列。 */
export function applyCommaStyleFromFormat(code: string | undefined): string {
  const raw = code?.trim() ?? "";
  if (raw === "" || raw === "General") {
    return "#,##0.00";
  }
  const first = (raw.split(";")[0] ?? "").trim();
  const { body, hasPct } = stripPercentSuffix(first);
  const decimals = countFractionDigitsInPattern(body);
  const pct = hasPct ? "%" : "";
  if (decimals <= 0) {
    return `#,##0${pct}`;
  }
  const zeros = "0".repeat(decimals);
  return `#,##0.${zeros}${pct}`;
}

function formatNumberScientific(n: number, fractionDigits: number): string {
  const s = n.toExponential(fractionDigits);
  return s.replace(/e/i, "E").replace(/E\+/, "E+").replace(/E-/, "E-");
}

function formatWithCommaPattern(n: number, fractionDigits: number, useGrouping: boolean): string {
  const opts: Intl.NumberFormatOptions = {
    useGrouping,
    minimumFractionDigits: fractionDigits,
    maximumFractionDigits: fractionDigits,
  };
  return new Intl.NumberFormat("zh-CN", opts).format(n);
}

function inferFractionDigitsFromPattern(part: string): number {
  const dot = part.indexOf(".");
  if (dot < 0) {
    return 0;
  }
  let c = 0;
  for (let i = dot + 1; i < part.length; i++) {
    if (part[i] === "0" || part[i] === "#") {
      c++;
    } else {
      break;
    }
  }
  return c;
}

/** 将数值按 Excel 格式串显示（子集：常规、百分比、千分位、货币符号、科学计数、日期/时间、文本 @）。 */
export function formatNumberWithExcelCode(value: number, formatCode: string): string {
  const fmt = formatCode.trim();
  if (fmt === "" || fmt === "General") {
    return String(value);
  }
  if (fmt === "@") {
    return String(value);
  }

  const lower = fmt.toLowerCase();
  if (lower.includes("e") && (lower.includes("0.00e") || lower.includes("0e"))) {
    const fd = inferFractionDigitsFromPattern(fmt.split(";")[0] ?? fmt);
    return formatNumberScientific(value, Math.max(2, fd));
  }

  if (fmt.includes("%")) {
    const first = (fmt.split(";")[0] ?? fmt).trim();
    const beforePct = first.endsWith("%") ? first.slice(0, -1) : first;
    const fd = inferFractionDigitsFromPattern(beforePct);
    const useComma = beforePct.includes(",");
    const scaled = value * 100;
    const opts: Intl.NumberFormatOptions = {
      minimumFractionDigits: fd,
      maximumFractionDigits: fd,
      useGrouping: useComma,
    };
    return `${new Intl.NumberFormat("zh-CN", opts).format(scaled)}%`;
  }

  if (fmt.includes("¥") || fmt.includes("$") || fmt.includes("￥")) {
    const first = (fmt.split(";")[0] ?? fmt).trim();
    const fd = inferFractionDigitsFromPattern(first.replace(/[¥$￥,]/g, ""));
    const sym = fmt.includes("¥") || fmt.includes("￥") ? "¥" : "$";
    const opts: Intl.NumberFormatOptions = {
      style: "currency",
      currency: sym === "¥" ? "CNY" : "USD",
      minimumFractionDigits: fd,
      maximumFractionDigits: fd,
    };
    try {
      return new Intl.NumberFormat("zh-CN", opts).format(value);
    } catch {
      return `${sym}${formatWithCommaPattern(value, fd, true)}`;
    }
  }

  if (/[yYmMdDhHsS]/.test(fmt) || fmt.includes("年")) {
    if (value > 0 && value < 1000000 && Number.isFinite(value)) {
      return formatDateLikeExcel(value, fmt);
    }
  }

  if (fmt.includes("#") && fmt.includes("?") && fmt.includes("/")) {
    return approximateFraction(value);
  }

  const first = (fmt.split(";")[0] ?? fmt).trim();
  const useComma = first.includes(",");
  const fd = inferFractionDigitsFromPattern(first.replace(/,/g, ""));
  return formatWithCommaPattern(value, fd, useComma);
}

function approximateFraction(x: number): string {
  if (!Number.isFinite(x)) {
    return String(x);
  }
  const sign = x < 0 ? "-" : "";
  const v = Math.abs(x);
  const maxD = 400;
  let bestA = 1;
  let bestB = 1;
  let bestErr = Math.abs(v - 1);
  for (let b = 1; b <= maxD; b++) {
    const a = Math.round(v * b);
    const err = Math.abs(v - a / b);
    if (err < bestErr) {
      bestErr = err;
      bestA = a;
      bestB = b;
    }
  }
  const g = gcd(bestA, bestB);
  return `${sign}${bestA / g}/${bestB / g}`;
}

function gcd(a: number, b: number): number {
  let x = Math.abs(a);
  let y = Math.abs(b);
  while (y !== 0) {
    const t = x % y;
    x = y;
    y = t;
  }
  return x || 1;
}

/**
 * 单元格显示文本：布尔与字符串保持原样；数字应用 `style.numberFormat`（若存在）。
 */
export function formatCellDisplayWithStyle(value: CellScalar, style: CellStyle | null | undefined): string {
  if (value === null || value === "") {
    return "";
  }
  if (typeof value === "boolean") {
    return value ? "TRUE" : "FALSE";
  }
  if (typeof value === "string") {
    return value;
  }
  if (typeof value === "number") {
    const nf = style?.numberFormat?.trim();
    if (nf === undefined || nf === "" || nf === "General") {
      return String(value);
    }
    try {
      return formatNumberWithExcelCode(value, nf);
    } catch {
      return String(value);
    }
  }
  return String(value);
}
