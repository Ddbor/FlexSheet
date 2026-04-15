/**
 * 条件格式 → OOXML（`conditionalFormatting` + `styles.xml` / `dxfs`），供 Excel 打开后正确显示。
 */

import {
  cfFormatPresetToOverlay,
  type CellBorderSide,
  type CfColorScaleEndpoint,
  type CfDataBarMaxEndpointType,
  type CfDataBarMinEndpointType,
  type CfDateOccurring,
  type CfIconSetId,
  type CfTextOperator,
  type CfValueOperator,
  type ConditionalFormatRule,
  type ConditionalFormattingOverlay,
  type Workbook,
  normalizeSelectionRange,
} from "@flexsheet/core";
import { formatCellRef } from "./a1.js";
import { escapeXml, sanitizeXml10Text } from "./xml-escape.js";

function rangeToSqref(range: ConditionalFormatRule["range"]): string {
  const n = normalizeSelectionRange(range);
  return `${formatCellRef(n.startRow, n.startCol)}:${formatCellRef(n.endRow, n.endCol)}`;
}

function anchorRef(range: ConditionalFormatRule["range"]): string {
  const n = normalizeSelectionRange(range);
  return formatCellRef(n.startRow, n.startCol);
}

function escapeExcelQuotedString(s: string): string {
  return s.replace(/"/g, '""');
}

function dxfBorderSideXml(side: CellBorderSide | undefined, tag: "left" | "right" | "top" | "bottom"): string {
  if (side === undefined) {
    return `<${tag}/>`;
  }
  const rgb =
    side.colorArgb !== undefined && side.colorArgb !== ""
      ? escapeXml(side.colorArgb)
      : "FF000000";
  return `<${tag} style="${escapeXml(side.kind)}"><color rgb="${rgb}"/></${tag}>`;
}

/**
 * 将条件格式叠加样式写成 `dxf`（仅包含有值的子节点，与 Excel 子集兼容）。
 */
export function conditionalFormatOverlayToDxfXml(overlay: ConditionalFormattingOverlay): string | null {
  const hasFont = overlay.fgArgb !== undefined && overlay.fgArgb !== "";
  const hasFill = overlay.fillArgb !== undefined && overlay.fillArgb !== "";
  const hasBorder =
    overlay.borderTop !== undefined ||
    overlay.borderLeft !== undefined ||
    overlay.borderBottom !== undefined ||
    overlay.borderRight !== undefined;
  if (!hasFont && !hasFill && !hasBorder) {
    return null;
  }
  const parts: string[] = [];
  if (hasFont) {
    parts.push(`<font><color rgb="${escapeXml(overlay.fgArgb as string)}"/></font>`);
  }
  if (hasFill) {
    parts.push(
      `<fill><patternFill patternType="solid"><fgColor rgb="${escapeXml(overlay.fillArgb as string)}"/><bgColor indexed="64"/></patternFill></fill>`,
    );
  }
  if (hasBorder) {
    parts.push(
      `<border>${dxfBorderSideXml(overlay.borderLeft, "left")}${dxfBorderSideXml(
        overlay.borderRight,
        "right",
      )}${dxfBorderSideXml(overlay.borderTop, "top")}${dxfBorderSideXml(overlay.borderBottom, "bottom")}<diagonal/></border>`,
    );
  }
  return `<dxf>${parts.join("")}</dxf>`;
}

function classicRuleDxfXml(rule: ConditionalFormatRule): string | null {
  if (rule.uiFamily !== "classic") {
    return null;
  }
  const overlay = cfFormatPresetToOverlay(rule.formatPreset, rule.customFormat);
  const dx = conditionalFormatOverlayToDxfXml(overlay);
  if (dx !== null) {
    return dx;
  }
  if (rule.classicType === "formula" && (rule.formulaExpression ?? "").trim() !== "") {
    return `<dxf/>`;
  }
  return null;
}

/** 全局 `dxfs` 顺序与 `dxfId` 索引（按工作簿表顺序 + 表内规则顺序）。 */
export interface WorkbookCfDxfIndex {
  readonly dxfXmlList: readonly string[];
  /** `${sheetIndex}:${ruleId}` → `dxfId` */
  readonly dxfIdBySheetAndRule: ReadonlyMap<string, number>;
}

export function buildWorkbookConditionalFormatDxfIndex(workbook: Workbook): WorkbookCfDxfIndex {
  const dxfXmlList: string[] = [];
  const dxfIdBySheetAndRule = new Map<string, number>();
  for (let si = 0; si < workbook.sheetCount; si++) {
    const sh = workbook.getSheet(si);
    if (sh === undefined) {
      continue;
    }
    for (const rule of sh.getConditionalFormatRules()) {
      const xml = classicRuleDxfXml(rule);
      if (xml === null) {
        continue;
      }
      const key = `${si}:${rule.id}`;
      if (dxfIdBySheetAndRule.has(key)) {
        continue;
      }
      dxfIdBySheetAndRule.set(key, dxfXmlList.length);
      dxfXmlList.push(xml);
    }
  }
  return { dxfXmlList, dxfIdBySheetAndRule };
}

function dxfIdFor(
  index: WorkbookCfDxfIndex,
  sheetIndex: number,
  rule: ConditionalFormatRule,
): number | undefined {
  return index.dxfIdBySheetAndRule.get(`${sheetIndex}:${rule.id}`);
}

function mapValueOpToOoxml(op: CfValueOperator): string {
  return op;
}

function mapTextOpToContainsTextOperator(op: CfTextOperator): string {
  switch (op) {
    case "contains":
      return "containsText";
    case "notContains":
      return "notContains";
    case "beginsWith":
      return "beginsWith";
    case "endsWith":
      return "endsWith";
    default:
      return "containsText";
  }
}

function colorScaleCfvoXml(ep: CfColorScaleEndpoint): string {
  const t = ep.type;
  const vRaw = ep.value.trim();
  if (t === "lowest") {
    return `<cfvo type="min"/>`;
  }
  if (t === "highest") {
    return `<cfvo type="max"/>`;
  }
  if (t === "number") {
    return `<cfvo type="num" val="${escapeXml(vRaw)}"/>`;
  }
  if (t === "percent") {
    return `<cfvo type="percent" val="${escapeXml(vRaw)}"/>`;
  }
  if (t === "percentile") {
    return `<cfvo type="percentile" val="${escapeXml(vRaw)}"/>`;
  }
  if (t === "formula") {
    const f = sanitizeXml10Text(vRaw.startsWith("=") ? vRaw.slice(1) : vRaw);
    return `<cfvo type="formula"><f>${f}</f></cfvo>`;
  }
  return `<cfvo type="min"/>`;
}

function colorRgbXml(argb: string): string {
  const s = argb.trim();
  if (s.length === 8) {
    return `<color rgb="${escapeXml(s)}"/>`;
  }
  if (s.length === 6) {
    return `<color rgb="FF${escapeXml(s)}"/>`;
  }
  return `<color rgb="FF000000"/>`;
}

function colorScaleRuleXml(rule: ConditionalFormatRule, priority: number): string | null {
  if (rule.uiFamily === "twoColorScale") {
    const min = rule.cfTwoColorMin;
    const max = rule.cfTwoColorMax;
    if (min === undefined || max === undefined) {
      return null;
    }
    const inner =
      `<colorScale>` +
      `${colorScaleCfvoXml(min)}${colorScaleCfvoXml(max)}` +
      `${colorRgbXml(min.colorArgb)}${colorRgbXml(max.colorArgb)}` +
      `</colorScale>`;
    return `<cfRule type="colorScale" priority="${priority}">${inner}</cfRule>`;
  }
  if (rule.uiFamily === "threeColorScale") {
    const min = rule.cfThreeColorMin;
    const mid = rule.cfThreeColorMid;
    const max = rule.cfThreeColorMax;
    if (min === undefined || mid === undefined || max === undefined) {
      return null;
    }
    const inner =
      `<colorScale>` +
      `${colorScaleCfvoXml(min)}${colorScaleCfvoXml(mid)}${colorScaleCfvoXml(max)}` +
      `${colorRgbXml(min.colorArgb)}${colorRgbXml(mid.colorArgb)}${colorRgbXml(max.colorArgb)}` +
      `</colorScale>`;
    return `<cfRule type="colorScale" priority="${priority}">${inner}</cfRule>`;
  }
  return null;
}

function dataBarMinCfvo(t: CfDataBarMinEndpointType, value: string): string {
  const v = value.trim();
  switch (t) {
    case "automatic":
      return `<cfvo type="autoMin"/>`;
    case "lowest":
      return `<cfvo type="min"/>`;
    case "number":
      return `<cfvo type="num" val="${escapeXml(v)}"/>`;
    case "percent":
      return `<cfvo type="percent" val="${escapeXml(v)}"/>`;
    case "percentile":
      return `<cfvo type="percentile" val="${escapeXml(v)}"/>`;
    case "formula": {
      const f = sanitizeXml10Text(v.startsWith("=") ? v.slice(1) : v);
      return `<cfvo type="formula"><f>${f}</f></cfvo>`;
    }
    default:
      return `<cfvo type="autoMin"/>`;
  }
}

function dataBarMaxCfvo(t: CfDataBarMaxEndpointType, value: string): string {
  const v = value.trim();
  switch (t) {
    case "automatic":
      return `<cfvo type="autoMax"/>`;
    case "highest":
      return `<cfvo type="max"/>`;
    case "number":
      return `<cfvo type="num" val="${escapeXml(v)}"/>`;
    case "percent":
      return `<cfvo type="percent" val="${escapeXml(v)}"/>`;
    case "percentile":
      return `<cfvo type="percentile" val="${escapeXml(v)}"/>`;
    case "formula": {
      const f = sanitizeXml10Text(v.startsWith("=") ? v.slice(1) : v);
      return `<cfvo type="formula"><f>${f}</f></cfvo>`;
    }
    default:
      return `<cfvo type="autoMax"/>`;
  }
}

function dataBarRuleXml(rule: ConditionalFormatRule, priority: number): string | null {
  if (rule.uiFamily !== "dataBar") {
    return null;
  }
  const minT = rule.cfDataBarMin?.type ?? "automatic";
  const minV = rule.cfDataBarMin?.value ?? "";
  const maxT = rule.cfDataBarMax?.type ?? "automatic";
  const maxV = rule.cfDataBarMax?.value ?? "";
  const pos = (rule.cfDataBarPositiveFillArgb ?? "FF638EC6").trim();
  const inner =
    `<dataBar>` +
    `${dataBarMinCfvo(minT, minV)}${dataBarMaxCfvo(maxT, maxV)}` +
    `${colorRgbXml(pos)}` +
    `</dataBar>`;
  return `<cfRule type="dataBar" priority="${priority}">${inner}</cfRule>`;
}

const ICON_SET_OOXML: Record<CfIconSetId, string> = {
  traffic3: "3TrafficLights1",
  arrows3: "3Arrows",
  arrows3_gray: "3ArrowsGray",
  flags3: "3Flags",
  shapes3: "3Symbols",
  signs3: "3Signs",
  stars3: "3Stars",
  bars4: "4Rating",
  bars5: "5Rating",
};

function iconSetRuleXml(rule: ConditionalFormatRule, priority: number): string | null {
  if (rule.uiFamily !== "iconSet") {
    return null;
  }
  const id = rule.cfIconSetId;
  if (id === undefined) {
    return null;
  }
  const setName = ICON_SET_OOXML[id];
  if (setName === undefined) {
    return null;
  }
  const entry = rule.cfIconThresholds;
  const count = setName.startsWith("5") ? 5 : setName.startsWith("4") ? 4 : 3;
  const cfvos: string[] = [];
  if (entry !== undefined && entry.length > 0) {
    for (const row of entry) {
      const vt = row.valueType;
      const val = row.value.trim();
      if (vt === "number") {
        cfvos.push(`<cfvo type="num" val="${escapeXml(val)}"/>`);
      } else if (vt === "percent") {
        cfvos.push(`<cfvo type="percent" val="${escapeXml(val)}"/>`);
      } else if (vt === "percentile") {
        cfvos.push(`<cfvo type="percentile" val="${escapeXml(val)}"/>`);
      } else {
        const f = sanitizeXml10Text(val.startsWith("=") ? val.slice(1) : val);
        cfvos.push(`<cfvo type="formula"><f>${f}</f></cfvo>`);
      }
    }
  } else {
    if (count === 3) {
      cfvos.push(`<cfvo type="percent" val="0"/>`);
      cfvos.push(`<cfvo type="percent" val="33"/>`);
      cfvos.push(`<cfvo type="percent" val="67"/>`);
    } else if (count === 4) {
      cfvos.push(`<cfvo type="percent" val="0"/>`);
      cfvos.push(`<cfvo type="percent" val="25"/>`);
      cfvos.push(`<cfvo type="percent" val="50"/>`);
      cfvos.push(`<cfvo type="percent" val="75"/>`);
    } else {
      cfvos.push(`<cfvo type="percent" val="0"/>`);
      cfvos.push(`<cfvo type="percent" val="20"/>`);
      cfvos.push(`<cfvo type="percent" val="40"/>`);
      cfvos.push(`<cfvo type="percent" val="60"/>`);
      cfvos.push(`<cfvo type="percent" val="80"/>`);
    }
  }
  const showVal = rule.cfIconShowIconOnly === true ? ` showValue="0"` : ` showValue="1"`;
  const rev = rule.cfIconReverseOrder === true ? ` reverse="1"` : ` reverse="0"`;
  const inner = `<iconSet iconSet="${escapeXml(setName)}"${showVal}${rev}>${cfvos.join("")}</iconSet>`;
  return `<cfRule type="iconSet" priority="${priority}">${inner}</cfRule>`;
}

/** 与 Excel「发生日期」规则等价的公式（锚格为选区左上角）。 */
function timePeriodFormula(a1: string, kind: CfDateOccurring): string {
  switch (kind) {
    case "today":
      return `INT(${a1})=INT(TODAY())`;
    case "yesterday":
      return `INT(${a1})=INT(TODAY()-1)`;
    case "tomorrow":
      return `INT(${a1})=INT(TODAY()+1)`;
    case "thisWeek":
      return `AND(INT(${a1})>=INT(TODAY()-WEEKDAY(TODAY(),2)+1),INT(${a1})<=INT(TODAY()-WEEKDAY(TODAY(),2)+7))`;
    case "lastWeek":
      return `AND(INT(${a1})>=INT(TODAY()-WEEKDAY(TODAY(),2)-6),INT(${a1})<=INT(TODAY()-WEEKDAY(TODAY(),2)))`;
    case "nextWeek":
      return `AND(INT(${a1})>=INT(TODAY()-WEEKDAY(TODAY(),2)+8),INT(${a1})<=INT(TODAY()-WEEKDAY(TODAY(),2)+14))`;
    case "thisMonth":
      return `AND(MONTH(${a1})=MONTH(TODAY()),YEAR(${a1})=YEAR(TODAY()))`;
    case "lastMonth":
      return `AND(MONTH(${a1})=MONTH(EDATE(TODAY(),-1)),YEAR(${a1})=YEAR(EDATE(TODAY(),-1)))`;
    case "nextMonth":
      return `AND(MONTH(${a1})=MONTH(EDATE(TODAY(),1)),YEAR(${a1})=YEAR(EDATE(TODAY(),1)))`;
    default:
      return `INT(${a1})=INT(TODAY())`;
  }
}

function dateTimePeriodAttribute(kind: NonNullable<ConditionalFormatRule["dateOccurring"]>): string {
  switch (kind) {
    case "today":
      return "today";
    case "yesterday":
      return "yesterday";
    case "tomorrow":
      return "tomorrow";
    case "thisWeek":
      return "thisWeek";
    case "lastWeek":
      return "lastWeek";
    case "nextWeek":
      return "nextWeek";
    case "thisMonth":
      return "thisMonth";
    case "lastMonth":
      return "lastMonth";
    case "nextMonth":
      return "nextMonth";
    default:
      return "today";
  }
}

function classicCfRuleXml(
  rule: ConditionalFormatRule,
  priority: number,
  dxfId: number | undefined,
): string | null {
  if (rule.uiFamily !== "classic") {
    return null;
  }
  const dxfAttr = dxfId !== undefined ? ` dxfId="${dxfId}"` : "";
  const a1 = anchorRef(rule.range);

  switch (rule.classicType) {
    case "formula": {
      const raw = (rule.formulaExpression ?? "").trim();
      if (raw === "") {
        return null;
      }
      const f = sanitizeXml10Text(raw.startsWith("=") ? raw.slice(1) : raw);
      if (dxfId === undefined) {
        return null;
      }
      return `<cfRule type="expression" priority="${priority}"${dxfAttr}><formula>${f}</formula></cfRule>`;
    }
    case "uniqueOrDuplicate": {
      const kind = rule.uniqueKind ?? "duplicate";
      if (dxfId === undefined) {
        return null;
      }
      const ty = kind === "unique" ? "uniqueValues" : "duplicateValues";
      return `<cfRule type="${ty}" priority="${priority}"${dxfAttr}/>`;
    }
    case "aboveBelowAverage": {
      if (dxfId === undefined) {
        return null;
      }
      const ak = rule.averageKind ?? "above";
      const aboveAverage = ak === "below" || ak === "equalOrBelow" ? "0" : "1";
      const equalAverage = ak === "equalOrAbove" || ak === "equalOrBelow" ? "1" : "0";
      return `<cfRule type="aboveAverage" priority="${priority}"${dxfAttr} aboveAverage="${aboveAverage}" equalAverage="${equalAverage}"/>`;
    }
    case "topBottomRanked": {
      if (dxfId === undefined) {
        return null;
      }
      const kind = rule.topBottomKind ?? "top";
      const n = Math.max(1, Math.min(1000, Math.round(rule.topBottomN ?? 10)));
      const bottom = kind === "bottom" || kind === "bottomPercent" ? "1" : "0";
      const percent = kind === "topPercent" || kind === "bottomPercent" ? "1" : "0";
      return `<cfRule type="top10" priority="${priority}"${dxfAttr} rank="${n}" percent="${percent}" bottom="${bottom}"/>`;
    }
    case "cellsThatContain": {
      const sub = rule.cellsThatContainKind ?? "cellValue";
      if (sub === "blanks") {
        if (dxfId === undefined) {
          return null;
        }
        return `<cfRule type="containsBlanks" priority="${priority}"${dxfAttr}><formula>LEN(TRIM(${a1}))=0</formula></cfRule>`;
      }
      if (sub === "noBlanks") {
        if (dxfId === undefined) {
          return null;
        }
        return `<cfRule type="notContainsBlanks" priority="${priority}"${dxfAttr}><formula>LEN(TRIM(${a1}))&gt;0</formula></cfRule>`;
      }
      if (sub === "errors") {
        if (dxfId === undefined) {
          return null;
        }
        return `<cfRule type="containsErrors" priority="${priority}"${dxfAttr}><formula>ISERROR(${a1})</formula></cfRule>`;
      }
      if (sub === "noErrors") {
        if (dxfId === undefined) {
          return null;
        }
        return `<cfRule type="notContainsErrors" priority="${priority}"${dxfAttr}><formula>NOT(ISERROR(${a1}))</formula></cfRule>`;
      }
      if (sub === "specificText") {
        if (dxfId === undefined) {
          return null;
        }
        const op = mapTextOpToContainsTextOperator(rule.textOperator ?? "contains");
        const needle = escapeXml(rule.value1 ?? "");
        const raw = (rule.value1 ?? "").trim();
        const q = escapeExcelQuotedString(raw);
        let fBody: string;
        if (op === "notContains") {
          fBody = `ISERROR(SEARCH("${q}",${a1}))`;
        } else if (op === "beginsWith") {
          fBody = `LEFT(${a1},LEN("${q}"))="${q}"`;
        } else if (op === "endsWith") {
          fBody = `RIGHT(${a1},LEN("${q}"))="${q}"`;
        } else {
          fBody = `NOT(ISERROR(SEARCH("${q}",${a1})))`;
        }
        const f = sanitizeXml10Text(fBody);
        return `<cfRule type="containsText" priority="${priority}"${dxfAttr} operator="${escapeXml(op)}" text="${needle}"><formula>${f}</formula></cfRule>`;
      }
      if (sub === "dateOccurring") {
        if (dxfId === undefined) {
          return null;
        }
        const tp = dateTimePeriodAttribute(rule.dateOccurring ?? "today");
        const fBody = timePeriodFormula(a1, rule.dateOccurring ?? "today");
        const f = sanitizeXml10Text(fBody);
        return `<cfRule type="timePeriod" priority="${priority}"${dxfAttr} timePeriod="${escapeXml(tp)}"><formula>${f}</formula></cfRule>`;
      }
      if (sub === "cellValue") {
        if (dxfId === undefined) {
          return null;
        }
        const op = mapValueOpToOoxml(rule.valueOperator ?? "greaterThan");
        const v1 = (rule.value1 ?? "").trim();
        const v2 = (rule.value2 ?? "").trim();
        if (op === "between" || op === "notBetween") {
          const f1 = sanitizeXml10Text(v1);
          const f2 = sanitizeXml10Text(v2);
          return `<cfRule type="cellIs" priority="${priority}"${dxfAttr} operator="${op}"><formula>${f1}</formula><formula>${f2}</formula></cfRule>`;
        }
        const f1 = sanitizeXml10Text(v1);
        return `<cfRule type="cellIs" priority="${priority}"${dxfAttr} operator="${op}"><formula>${f1}</formula></cfRule>`;
      }
      return null;
    }
    default:
      return null;
  }
}

function cfRuleXmlForRule(
  rule: ConditionalFormatRule,
  priority: number,
  dxfIndex: WorkbookCfDxfIndex,
  sheetIndex: number,
): string | null {
  const scale = colorScaleRuleXml(rule, priority);
  if (scale !== null) {
    return scale;
  }
  const db = dataBarRuleXml(rule, priority);
  if (db !== null) {
    return db;
  }
  const ic = iconSetRuleXml(rule, priority);
  if (ic !== null) {
    return ic;
  }
  const dxfId = dxfIdFor(dxfIndex, sheetIndex, rule);
  return classicCfRuleXml(rule, priority, dxfId);
}

/**
 * 生成工作表 `conditionalFormatting` 片段（不含根元素）；无规则时返回空串。
 */
export function buildSheetConditionalFormattingXml(
  sheetIndex: number,
  rules: readonly ConditionalFormatRule[],
  dxfIndex: WorkbookCfDxfIndex,
): string {
  if (rules.length === 0) {
    return "";
  }
  const blocks: string[] = [];
  let p = 1;
  for (const rule of rules) {
    const sq = rangeToSqref(rule.range);
    const inner = cfRuleXmlForRule(rule, p++, dxfIndex, sheetIndex);
    if (inner === null) {
      continue;
    }
    blocks.push(`<conditionalFormatting sqref="${escapeXml(sq)}">${inner}</conditionalFormatting>`);
  }
  return blocks.join("");
}

export function expandBoundsWithConditionalFormatRanges(
  sheet: { getConditionalFormatRules(): readonly ConditionalFormatRule[] },
  b: { minR: number; maxR: number; minC: number; maxC: number } | null,
): { minR: number; maxR: number; minC: number; maxC: number } | null {
  let cur = b;
  for (const rule of sheet.getConditionalFormatRules()) {
    const n = normalizeSelectionRange(rule.range);
    if (cur === null) {
      cur = { minR: n.startRow, maxR: n.endRow, minC: n.startCol, maxC: n.endCol };
    } else {
      cur = {
        minR: Math.min(cur.minR, n.startRow),
        maxR: Math.max(cur.maxR, n.endRow),
        minC: Math.min(cur.minC, n.startCol),
        maxC: Math.max(cur.maxC, n.endCol),
      };
    }
  }
  return cur;
}
