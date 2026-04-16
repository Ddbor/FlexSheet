import {
  adjustDecimalPlacesInFormat,
  applyCommaStyleFromFormat,
  normalizeSelectionRange,
  type SelectionRange,
  type Worksheet,
} from "@flexsheet/core";
import { RIBBON_FONT_FAMILY_ITEMS } from "./font-family-items.js";
import { argb8ToCssHex6, cssHexToFillArgb } from "./ribbon-color-argb.js";
import { showRibbonColorDialog } from "./ribbon-color-dialog.js";
import {
  showConditionalFormatManageRulesDialog,
  showConditionalFormatNewRuleDialog,
  type CfNewRuleDialogSeed,
} from "./ribbon-conditional-format-dialog.js";
import { applyRibbonCellStyleCommand } from "./cell-style-ribbon-handlers.js";
import { getNumberFormatPresetByCommandId } from "./ribbon-number-format-chrome.js";
import type { FlexSheetLike, RibbonCommandEvent } from "./ribbon-types.js";

async function openNewConditionalFormatRule(
  fs: FlexSheetLike,
  seed: CfNewRuleDialogSeed,
): Promise<void> {
  const range = fs.selection.getNormalizedRange();
  const rule = await showConditionalFormatNewRuleDialog(range, seed);
  if (rule !== null && fs.addConditionalFormatRuleFromUi !== undefined) {
    fs.addConditionalFormatRuleFromUi(rule);
  }
}

async function openManageConditionalFormatRules(fs: FlexSheetLike): Promise<void> {
  const sheet = fs.workbook?.getActiveSheet();
  if (sheet === undefined || fs.replaceConditionalFormatRulesFromUi === undefined) {
    return;
  }
  const next = await showConditionalFormatManageRulesDialog(
    sheet.getConditionalFormatRules(),
    sheet.name,
  );
  if (next !== null) {
    fs.replaceConditionalFormatRulesFromUi(next);
  }
}

const FONT_FAMILY_CSS = new Map<string, string>(
  RIBBON_FONT_FAMILY_ITEMS.map((it) => {
    const css =
      it.previewFontFamily ??
      (it.id === "home.font.family.wingdings" ? "Wingdings, fantasy" : `"${it.label}", sans-serif`);
    return [it.id, css];
  }),
);

function readPickHex(ev: RibbonCommandEvent): string | null {
  const h = ev.payload?.hex;
  if (typeof h !== "string") {
    return null;
  }
  const t = h.trim();
  if (!/^#[\dA-Fa-f]{3}$/i.test(t) && !/^#[\dA-Fa-f]{6}$/i.test(t)) {
    return null;
  }
  return t;
}

async function applyMoreColor(fs: FlexSheetLike, kind: "fill" | "font"): Promise<void> {
  const st = fs.getActiveCellStyle();
  const argb = kind === "fill" ? st?.fillArgb : st?.fgArgb;
  const argbTrim = typeof argb === "string" ? argb.trim() : "";
  const initial =
    argbTrim !== "" && /^[\dA-Fa-f]{8}$/i.test(argbTrim) ? argb8ToCssHex6(argbTrim) : "#ffffff";
  const picked = await showRibbonColorDialog(initial);
  if (picked === null) {
    return;
  }
  if (kind === "fill") {
    fs.applySelectionStylePatch({ fillArgb: cssHexToFillArgb(picked) });
  } else {
    fs.applySelectionStylePatch({ fgArgb: cssHexToFillArgb(picked) });
  }
}

function readRangeFontState(
  sheet: Worksheet,
  range: SelectionRange,
): {
  readonly allBold: boolean;
  readonly allItalic: boolean;
  readonly allSingleUnderline: boolean;
  readonly allDoubleUnderline: boolean;
} {
  const n = normalizeSelectionRange(range);
  let allBold = true;
  let allItalic = true;
  let allSingleUnderline = true;
  let allDoubleUnderline = true;
  for (let r = n.startRow; r <= n.endRow; r++) {
    for (let c = n.startCol; c <= n.endCol; c++) {
      const st = sheet.getCell(r, c).style;
      if (st?.bold !== true) {
        allBold = false;
      }
      if (st?.italic !== true) {
        allItalic = false;
      }
      if (st?.underline !== "single") {
        allSingleUnderline = false;
      }
      if (st?.underline !== "double") {
        allDoubleUnderline = false;
      }
    }
  }
  return { allBold, allItalic, allSingleUnderline, allDoubleUnderline };
}

function readRangeWrapState(sheet: Worksheet, range: SelectionRange): boolean {
  const n = normalizeSelectionRange(range);
  let anyCell = false;
  let allWrap = true;
  for (let r = n.startRow; r <= n.endRow; r++) {
    for (let c = n.startCol; c <= n.endCol; c++) {
      anyCell = true;
      if (sheet.getCell(r, c).style?.wrapText !== true) {
        allWrap = false;
      }
    }
  }
  return anyCell && allWrap;
}

const RIBBON_BORDER_COMMAND_IDS = new Set<string>([
  "home.font.border",
  "home.font.border.bottom",
  "home.font.border.top",
  "home.font.border.left",
  "home.font.border.right",
  "home.font.border.none",
  "home.font.border.all",
  "home.font.border.outside",
  "home.font.border.thickBox",
  "home.font.border.doubleBottom",
  "home.font.border.thickBottom",
  "home.font.border.topBottom",
  "home.font.border.topThickBottom",
  "home.font.border.topDoubleBottom",
]);

function tryApplyRibbonBorder(ev: RibbonCommandEvent, fs: FlexSheetLike): boolean {
  if (!RIBBON_BORDER_COMMAND_IDS.has(ev.id)) {
    return false;
  }
  if (fs.applyRibbonBorderCommand === undefined) {
    return false;
  }
  fs.applyRibbonBorderCommand(ev.id);
  return true;
}

/**
 * 将部分 Ribbon 命令映射到 FlexSheet（非「视图」选项卡逻辑可放此处）。
 * 「视图」选项卡由 `ViewRibbonController` 统一处理。
 */
export function applyRibbonCommandToFlexSheet(ev: RibbonCommandEvent, fs: FlexSheetLike): boolean {
  if (tryApplyRibbonBorder(ev, fs)) {
    return true;
  }
  if (applyRibbonCellStyleCommand(fs, ev.id)) {
    return true;
  }
  const preset = getNumberFormatPresetByCommandId(ev.id);
  if (preset !== undefined) {
    fs.applySelectionStylePatch({ numberFormat: preset.format === "" ? null : preset.format });
    return true;
  }
  switch (ev.id) {
    case "home.undo.back":
      fs.undo();
      return true;
    case "home.undo.forward":
      fs.redo();
      return true;
    case "home.clipboard.copy":
      void fs.clipboardCopy();
      return true;
    case "home.clipboard.cut":
      void fs.clipboardCut();
      return true;
    case "home.clipboard.paste":
      void fs.clipboardPaste();
      return true;
    case "home.font.grow":
      fs.applySelectionFontSizeStep(1);
      return true;
    case "home.font.shrink":
      fs.applySelectionFontSizeStep(-1);
      return true;
    case "home.font.fill.pick": {
      const hex = readPickHex(ev);
      if (hex === null) {
        return false;
      }
      fs.applySelectionStylePatch({ fillArgb: cssHexToFillArgb(hex) });
      return true;
    }
    case "home.font.fill.none":
      fs.applySelectionStylePatch({ fillArgb: null });
      return true;
    case "home.font.fill.more":
      void applyMoreColor(fs, "fill");
      return true;
    case "home.font.color.pick": {
      const hex = readPickHex(ev);
      if (hex === null) {
        return false;
      }
      fs.applySelectionStylePatch({ fgArgb: cssHexToFillArgb(hex) });
      return true;
    }
    case "home.font.color.none":
      fs.applySelectionStylePatch({ fgArgb: null });
      return true;
    case "home.font.color.more":
      void applyMoreColor(fs, "font");
      return true;
    case "home.font.bold": {
      const sheet = fs.workbook?.getActiveSheet();
      if (sheet === undefined) {
        return false;
      }
      const { allBold } = readRangeFontState(sheet, fs.selection.getNormalizedRange());
      fs.applySelectionStylePatch({ bold: !allBold });
      return true;
    }
    case "home.font.italic": {
      const sheet = fs.workbook?.getActiveSheet();
      if (sheet === undefined) {
        return false;
      }
      const { allItalic } = readRangeFontState(sheet, fs.selection.getNormalizedRange());
      fs.applySelectionStylePatch({ italic: !allItalic });
      return true;
    }
    case "home.font.underline": {
      const sheet = fs.workbook?.getActiveSheet();
      if (sheet === undefined) {
        return false;
      }
      const { allSingleUnderline } = readRangeFontState(sheet, fs.selection.getNormalizedRange());
      fs.applySelectionStylePatch(
        allSingleUnderline ? { underline: null } : { underline: "single" },
      );
      return true;
    }
    case "home.font.doubleUnderline": {
      const sheet = fs.workbook?.getActiveSheet();
      if (sheet === undefined) {
        return false;
      }
      const { allDoubleUnderline } = readRangeFontState(sheet, fs.selection.getNormalizedRange());
      fs.applySelectionStylePatch(
        allDoubleUnderline ? { underline: null } : { underline: "double" },
      );
      return true;
    }
    case "home.number.quick.percent":
      fs.applySelectionStylePatch({ numberFormat: "0%" });
      return true;
    case "home.number.quick.comma": {
      const cur = fs.getActiveCellStyle()?.numberFormat;
      fs.applySelectionStylePatch({ numberFormat: applyCommaStyleFromFormat(cur) });
      return true;
    }
    case "home.number.quick.increaseDecimal": {
      const cur = fs.getActiveCellStyle()?.numberFormat;
      const next = adjustDecimalPlacesInFormat(cur, 1);
      fs.applySelectionStylePatch({ numberFormat: next === "" ? null : next });
      return true;
    }
    case "home.number.quick.decreaseDecimal": {
      const cur = fs.getActiveCellStyle()?.numberFormat;
      const next = adjustDecimalPlacesInFormat(cur, -1);
      fs.applySelectionStylePatch({ numberFormat: next === "" ? null : next });
      return true;
    }
    case "home.cells.fill.down":
    case "home.cells.fill.right":
    case "home.cells.fill.up":
    case "home.cells.fill.left": {
      if (fs.applySelectionFillDirection === undefined) {
        return false;
      }
      const dir = ev.id.slice("home.cells.fill.".length) as "down" | "right" | "up" | "left";
      fs.applySelectionFillDirection(dir);
      return true;
    }
    case "home.cells.fill.series": {
      if (fs.openFillSeriesDialog === undefined) {
        return false;
      }
      fs.openFillSeriesDialog();
      return true;
    }
    case "home.cells.clear.contents": {
      if (fs.clearSelectionContents === undefined) {
        return false;
      }
      fs.clearSelectionContents();
      return true;
    }
    case "home.cells.clear.formats": {
      if (fs.clearSelectionFormats === undefined) {
        return false;
      }
      fs.clearSelectionFormats();
      return true;
    }
    case "home.cells.clear.all": {
      if (fs.clearSelectionAll === undefined) {
        return false;
      }
      fs.clearSelectionAll();
      return true;
    }
    case "home.align.top":
      fs.applySelectionStylePatch({ vAlign: "top" });
      return true;
    case "home.align.middle":
      fs.applySelectionStylePatch({ vAlign: "middle" });
      return true;
    case "home.align.bottom":
      fs.applySelectionStylePatch({ vAlign: "bottom" });
      return true;
    case "home.align.left":
      fs.applySelectionStylePatch({ hAlign: "left", indentLevel: null });
      return true;
    case "home.align.center":
      fs.applySelectionStylePatch({ hAlign: "center", indentLevel: null });
      return true;
    case "home.align.right":
      fs.applySelectionStylePatch({ hAlign: "right", indentLevel: null });
      return true;
    case "home.align.indentIncrease":
      fs.applySelectionIndentStep(1);
      return true;
    case "home.align.indentDecrease":
      fs.applySelectionIndentStep(-1);
      return true;
    case "home.align.wrap": {
      const sheet = fs.workbook?.getActiveSheet();
      if (sheet === undefined) {
        return false;
      }
      const allWrap = readRangeWrapState(sheet, fs.selection.getNormalizedRange());
      fs.applySelectionStylePatch({ wrapText: !allWrap });
      return true;
    }
    case "home.align.merge": {
      if (fs.applySelectionMerge === undefined) {
        return false;
      }
      fs.applySelectionMerge("mergeCenter");
      return true;
    }
    case "home.align.mergeAcross": {
      if (fs.applySelectionMerge === undefined) {
        return false;
      }
      fs.applySelectionMerge("mergeAcross");
      return true;
    }
    case "home.align.mergeCells": {
      if (fs.applySelectionMerge === undefined) {
        return false;
      }
      fs.applySelectionMerge("mergeCells");
      return true;
    }
    case "home.align.unmerge": {
      if (fs.applySelectionMerge === undefined) {
        return false;
      }
      fs.applySelectionMerge("unmerge");
      return true;
    }
    case "home.align.textDirection.counterClockwise":
      fs.applySelectionStylePatch({ textOrientation: "angleUp45" });
      return true;
    case "home.align.textDirection.clockwise":
      fs.applySelectionStylePatch({ textOrientation: "angleDown45" });
      return true;
    case "home.align.textDirection.vertical":
      fs.applySelectionStylePatch({ textOrientation: "verticalStack" });
      return true;
    case "home.align.textDirection.rotateUp":
      fs.applySelectionStylePatch({ textOrientation: "rotateUp90" });
      return true;
    case "home.align.textDirection.rotateDown":
      fs.applySelectionStylePatch({ textOrientation: "rotateDown90" });
      return true;
    case "data.sort.asc":
    case "data.sort.desc": {
      if (fs.sortSelectionRowsByKeyColumn === undefined) {
        return false;
      }
      const ac = fs.selection.getActiveCell();
      fs.sortSelectionRowsByKeyColumn(ac.col, {
        type: "value",
        direction: ev.id === "data.sort.asc" ? "asc" : "desc",
      });
      return true;
    }
    case "data.sort.custom": {
      if (fs.openCustomSortDialog === undefined) {
        return false;
      }
      fs.openCustomSortDialog();
      return true;
    }
    case "home.style.conditional.newRule":
    case "home.style.conditional.highlightCells.moreRules":
    case "home.style.conditional.topBottom.moreRules":
    case "home.style.conditional.dataBars.moreRules":
    case "home.style.conditional.colorScales.moreRules": {
      if (fs.addConditionalFormatRuleFromUi === undefined) {
        return false;
      }
      void openNewConditionalFormatRule(fs, { kind: "default" });
      return true;
    }
    case "home.style.conditional.manageRules": {
      if (fs.replaceConditionalFormatRulesFromUi === undefined || fs.workbook === undefined) {
        return false;
      }
      void openManageConditionalFormatRules(fs);
      return true;
    }
    case "home.style.conditional.clearRulesFromSelection": {
      if (fs.clearConditionalFormatRulesInSelection === undefined) {
        return false;
      }
      fs.clearConditionalFormatRulesInSelection();
      return true;
    }
    case "home.style.conditional.clearRulesFromSheet": {
      if (fs.clearAllConditionalFormatRulesFromUi === undefined) {
        return false;
      }
      fs.clearAllConditionalFormatRulesFromUi();
      return true;
    }
    case "home.style.table.newStyle": {
      if (fs.openNewTableStyleDialog === undefined) {
        return false;
      }
      fs.openNewTableStyleDialog();
      return true;
    }
    case "insert.pivottable.options": {
      if (fs.openPivotTableDialog === undefined) {
        return false;
      }
      fs.openPivotTableDialog();
      return true;
    }
    case "data.pivot.fields": {
      if (fs.openPivotTableFieldsPane === undefined) {
        return false;
      }
      fs.openPivotTableFieldsPane();
      return true;
    }
    default: {
      if (
        ev.id.startsWith("home.style.conditional.highlightCells.") &&
        ev.id !== "home.style.conditional.highlightCells" &&
        fs.addConditionalFormatRuleFromUi !== undefined
      ) {
        void openNewConditionalFormatRule(fs, {
          kind: "highlightPreset",
          highlightCommandId: ev.id,
        });
        return true;
      }
      if (
        ev.id.startsWith("home.style.conditional.topBottom.") &&
        ev.id !== "home.style.conditional.topBottom" &&
        fs.addConditionalFormatRuleFromUi !== undefined
      ) {
        void openNewConditionalFormatRule(fs, {
          kind: "topBottomPreset",
          topBottomCommandId: ev.id,
        });
        return true;
      }
      if (
        ev.id.startsWith("home.style.conditional.dataBars.") &&
        ev.id !== "home.style.conditional.dataBars" &&
        fs.addConditionalFormatRuleFromUi !== undefined
      ) {
        void openNewConditionalFormatRule(fs, {
          kind: "dataBarPreset",
          dataBarCommandId: ev.id,
        });
        return true;
      }
      if (
        ev.id.startsWith("home.style.conditional.colorScales.") &&
        ev.id !== "home.style.conditional.colorScales" &&
        fs.addConditionalFormatRuleFromUi !== undefined
      ) {
        void openNewConditionalFormatRule(fs, {
          kind: "colorScalePreset",
          colorScaleCommandId: ev.id,
        });
        return true;
      }
      const fam = FONT_FAMILY_CSS.get(ev.id);
      if (fam !== undefined) {
        fs.applySelectionStylePatch({ fontFamily: fam });
        return true;
      }
      const m = /^home\.font\.size\.(\d+)$/.exec(ev.id);
      if (m !== null) {
        const pt = Number(m[1]);
        if (Number.isFinite(pt) && pt > 0 && pt <= 409) {
          fs.applySelectionStylePatch({ fontSizePt: pt });
          return true;
        }
      }
      if (ev.id.startsWith("home.style.table.") && fs.openFormatAsTableFromRibbon !== undefined) {
        fs.openFormatAsTableFromRibbon(ev.id);
        return true;
      }
      return false;
    }
  }
}
