import type { CellScalar, CellStyle } from "./cell.js";

/** Excel 序列日期转 UTC `Date`（与 Excel 1900 日期系一致；`25569` 为 1970-01-01）。 */
export function excelSerialToUtcDate(serial: number): Date {
  return new Date(Math.round((serial - 25569) * 86400 * 1000));
}

function pad2(n: number): string {
  return n < 10 ? `0${n}` : String(n);
}

function formatDateLikeExcel(value: number, format: string): string {
  let fmt = format.trim();
  while (fmt.startsWith("*")) {
    fmt = fmt.slice(1).trim();
  }
  fmt = fmt.replace(/^[a-z]\s+/i, "").trim();
  const d = excelSerialToUtcDate(value);
  const y = d.getUTCFullYear();
  const m = d.getUTCMonth() + 1;
  const day = d.getUTCDate();
  const h = d.getUTCHours();
  const min = d.getUTCMinutes();
  const s = d.getUTCSeconds();

  if (/[hH]:/.test(fmt) || fmt.includes("秒") || fmt.includes("mm:ss")) {
    const hh = pad2(h);
    const mm = pad2(min);
    const ss = pad2(s);
    if (fmt.includes("秒") || fmt.split(":").length >= 3) {
      return `${hh}:${mm}:${ss}`;
    }
    return `${hh}:${mm}`;
  }

  if (fmt.includes("年") && fmt.includes("月")) {
    return `${y}年${m}月${day}日`;
  }
  if (/y{2,4}.*m{1,2}.*d{1,2}/i.test(fmt) || fmt.includes("短")) {
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

function splitExcelFormatSections(formatCode: string): string[] {
  return formatCode.split(";").map((s) => s.trim());
}

function pickExcelNumberFormatSection(formatCode: string, value: number): string {
  const parts = splitExcelFormatSections(formatCode);
  if (parts.length === 1) {
    return parts[0] ?? "";
  }
  if (value > 0) {
    return parts[0] ?? "";
  }
  if (value < 0) {
    return parts.length >= 2 ? (parts[1] ?? parts[0] ?? "") : (parts[0] ?? "");
  }
  return parts.length >= 3 ? (parts[2] ?? parts[0] ?? "") : (parts[0] ?? "");
}

/** 剥掉段首的 `[...]`（颜色、条件、地区等），便于解析数字/货币主体。 */
function stripExcelSectionLeadingTokens(section: string): string {
  let s = section.trim();
  for (;;) {
    const m = s.match(/^\[[^\]]+\]/);
    if (m === null) {
      break;
    }
    s = s.slice(m[0].length).trim();
  }
  return s;
}

/**
 * 多段格式下：负数括号段、无减号红色段等 → 调整参与格式化的数值与模式。
 */
function resolveSignedSectionPattern(
  value: number,
  formatCode: string,
): { readonly v: number; readonly pattern: string; readonly wrapParens: boolean } {
  const parts = splitExcelFormatSections(formatCode);
  const raw = pickExcelNumberFormatSection(formatCode, value);
  const stripped = stripExcelSectionLeadingTokens(raw);
  if (value < 0 && stripped.startsWith("(") && stripped.endsWith(")")) {
    return { v: Math.abs(value), pattern: stripped.slice(1, -1), wrapParens: true };
  }
  if (
    value < 0 &&
    parts.length >= 2 &&
    !stripped.startsWith("-") &&
    !stripped.startsWith("(")
  ) {
    return { v: Math.abs(value), pattern: stripped, wrapParens: false };
  }
  return { v: value, pattern: stripped, wrapParens: false };
}

function formatPercentSection(value: number, pattern: string): string {
  const first = pattern.trim();
  const { body, hasPct } = stripPercentSuffix(first);
  if (!hasPct) {
    return formatWithCommaPattern(value, inferFractionDigitsFromPattern(body.replace(/,/g, "")), body.includes(","));
  }
  const beforePct = body;
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

function formatCurrencySection(value: number, pattern: string): string {
  const first = pattern.trim();
  const fd = inferFractionDigitsFromPattern(first.replace(/[¥$￥,]/g, ""));
  const sym = first.includes("¥") || first.includes("￥") ? "¥" : "$";
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

/** 将数值按 Excel 格式串显示（子集：常规、百分比、千分位、货币符号、科学计数、日期/时间、文本 @；含简单多段正/负/零）。 */
export function formatNumberWithExcelCode(value: number, formatCode: string): string {
  const fmt = formatCode.trim();
  if (fmt === "" || fmt === "General") {
    return String(value);
  }
  if (fmt === "@") {
    return String(value);
  }

  if (/DBNum1/i.test(fmt)) {
    return formatChineseLowerInteger(value);
  }
  if (/DBNum2/i.test(fmt)) {
    return formatChineseUpperInteger(value);
  }

  const firstSection = (splitExcelFormatSections(fmt)[0] ?? "").trim();
  const lowerFirst = firstSection.toLowerCase();
  if (lowerFirst.includes("e") && (lowerFirst.includes("0.00e") || lowerFirst.includes("0e"))) {
    const fd = inferFractionDigitsFromPattern(firstSection);
    return formatNumberScientific(value, Math.max(2, fd));
  }

  if (fmt.includes("%")) {
    const { v, pattern, wrapParens } = resolveSignedSectionPattern(value, fmt);
    const inner = formatPercentSection(v, pattern);
    return wrapParens ? `(${inner})` : inner;
  }

  const secForDate = stripExcelSectionLeadingTokens(pickExcelNumberFormatSection(fmt, value));
  if (/[yYmMdDhHsS]/.test(secForDate) || secForDate.includes("年") || secForDate.includes("时")) {
    if (value > 0 && value < 1000000 && Number.isFinite(value)) {
      return formatDateLikeExcel(value, secForDate);
    }
  }

  if (fmt.includes("#") && fmt.includes("?") && fmt.includes("/")) {
    return approximateFraction(value);
  }

  const { v, pattern, wrapParens } = resolveSignedSectionPattern(value, fmt);
  const symPart = pattern;
  if (symPart.includes("¥") || symPart.includes("$") || symPart.includes("￥")) {
    const inner = formatCurrencySection(v, symPart);
    return wrapParens ? `(${inner})` : inner;
  }

  const useComma = symPart.includes(",");
  const fd = inferFractionDigitsFromPattern(symPart.replace(/,/g, ""));
  const inner = formatWithCommaPattern(v, fd, useComma);
  return wrapParens ? `(${inner})` : inner;
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

const ZH_LOWER_DIGITS = ["零", "一", "二", "三", "四", "五", "六", "七", "八", "九"] as const;
const ZH_UPPER_DIGITS = ["零", "壹", "贰", "叁", "肆", "伍", "陆", "柒", "捌", "玖"] as const;
const ZH_LOWER_UNITS = ["", "十", "百", "千"] as const;
const ZH_UPPER_UNITS = ["", "拾", "佰", "仟"] as const;

function formatFourDigitsChinese(
  k: number,
  digits: readonly string[],
  units: readonly string[],
  embedded: boolean,
): string {
  if (k === 0) {
    return "";
  }
  if (k < 10) {
    return digits[k]!;
  }
  if (k < 20) {
    if (!embedded && k === 10) {
      return units[1];
    }
    if (!embedded) {
      return `${units[1]}${k > 10 ? digits[k % 10]! : ""}`;
    }
    return `${digits[1]!}${units[1]}${k > 10 ? digits[k % 10]! : ""}`;
  }
  if (k < 100) {
    const tens = Math.floor(k / 10);
    const ones = k % 10;
    return `${digits[tens]!}${units[1]}${ones > 0 ? digits[ones]! : ""}`;
  }
  if (k < 1000) {
    const h = Math.floor(k / 100);
    const rest = k % 100;
    let out = `${digits[h]!}${units[2]}`;
    if (rest === 0) {
      return out;
    }
    if (rest < 10) {
      return `${out}${digits[0]}${digits[rest]!}`;
    }
    return `${out}${formatFourDigitsChinese(rest, digits, units, true)}`;
  }
  const t = Math.floor(k / 1000);
  const rest = k % 1000;
  let out = `${digits[t]!}${units[3]}`;
  if (rest === 0) {
    return out;
  }
  if (rest < 100) {
    return `${out}${digits[0]}${formatFourDigitsChinese(rest, digits, units, true)}`;
  }
  return `${out}${formatFourDigitsChinese(rest, digits, units, true)}`;
}

function formatChineseIntegerWithTables(
  value: number,
  digits: readonly string[],
  units: readonly string[],
): string {
  const sign = value < 0 ? "负" : "";
  const n = Math.trunc(Math.abs(value));
  if (n === 0) {
    return `${sign}${digits[0]}`;
  }
  const yi = Math.floor(n / 100000000);
  const rest1 = n % 100000000;
  const wan = Math.floor(rest1 / 10000);
  const low = rest1 % 10000;
  let s = "";
  if (yi > 0) {
    s += `${formatFourDigitsChinese(yi, digits, units, false)}亿`;
  }
  if (wan > 0) {
    if (yi > 0 && wan < 1000) {
      s += digits[0];
    }
    s += `${formatFourDigitsChinese(wan, digits, units, false)}万`;
  } else if (yi > 0 && low > 0) {
    s += digits[0];
  }
  if (low > 0) {
    if (wan > 0 && low < 1000) {
      s += digits[0];
    }
    s += formatFourDigitsChinese(low, digits, units, yi > 0 || wan > 0);
  }
  if (s === "") {
    s = digits[0]!;
  }
  return `${sign}${s}`;
}

function formatChineseLowerInteger(value: number): string {
  return formatChineseIntegerWithTables(value, ZH_LOWER_DIGITS, ZH_LOWER_UNITS);
}

function formatChineseUpperInteger(value: number): string {
  return formatChineseIntegerWithTables(value, ZH_UPPER_DIGITS, ZH_UPPER_UNITS);
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
