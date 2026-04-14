/**
 * Ribbon「新建格式规则 / 管理规则」：与 Excel 经典条件格式对话框布局相近。
 */

import type {
  CfAverageKind,
  CfCellsThatContainKind,
  CfColorScaleEndpointType,
  CfDataBarAxisPosition,
  CfDataBarBorderKind,
  CfDataBarDirection,
  CfDataBarFillKind,
  CfDataBarMaxEndpointType,
  CfDataBarMinEndpointType,
  CfDateOccurring,
  CfFormatPresetId,
  CfTextOperator,
  CfTopBottomKind,
  CfUniqueKind,
  CfValueOperator,
  ConditionalFormatClassicRuleType,
  ConditionalFormatRule,
  ConditionalFormatUiFamily,
  SelectionRange,
} from "@flexsheet/core";
import { cfFormatPresetToOverlay, type CellStylePatch } from "@flexsheet/core";
import { appendRibbonColorPaletteContent } from "./ribbon-color-picker-menu.js";
import { cssHexToFillArgb } from "./ribbon-color-argb.js";
import { showRibbonColorDialog } from "./ribbon-color-dialog.js";

function ribbonIsDarkTheme(): boolean {
  const r = document.querySelector(".fs-ribbon");
  return r?.getAttribute("data-theme") === "dark";
}

function newCfRuleId(): string {
  return `cf-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 9)}`;
}

const PREVIEW_TEXT = "微软卓越 AaBbCc";

const CF_MIN_TYPE_OPTS: { readonly v: CfColorScaleEndpointType; readonly t: string }[] = [
  { v: "lowest", t: "最低值" },
  { v: "number", t: "数字" },
  { v: "percent", t: "百分比" },
  { v: "formula", t: "公式" },
  { v: "percentile", t: "百分点值" },
];

const CF_MAX_TYPE_OPTS: { readonly v: CfColorScaleEndpointType; readonly t: string }[] = [
  { v: "highest", t: "最高值" },
  { v: "number", t: "数字" },
  { v: "percent", t: "百分比" },
  { v: "formula", t: "公式" },
  { v: "percentile", t: "百分点值" },
];

/** 三色刻度「中间值」：无最低/最高类型（与 Excel 一致）。 */
const CF_MID_TYPE_OPTS: { readonly v: CfColorScaleEndpointType; readonly t: string }[] = [
  { v: "number", t: "数字" },
  { v: "percent", t: "百分比" },
  { v: "formula", t: "公式" },
  { v: "percentile", t: "百分点值" },
];

const CF_DATA_BAR_MIN_OPTS: { readonly v: CfDataBarMinEndpointType; readonly t: string }[] = [
  { v: "lowest", t: "最低值" },
  { v: "number", t: "数字" },
  { v: "percent", t: "百分比" },
  { v: "formula", t: "公式" },
  { v: "percentile", t: "百分点值" },
  { v: "automatic", t: "自动" },
];

const CF_DATA_BAR_MAX_OPTS: { readonly v: CfDataBarMaxEndpointType; readonly t: string }[] = [
  { v: "highest", t: "最高值" },
  { v: "number", t: "数字" },
  { v: "percent", t: "百分比" },
  { v: "formula", t: "公式" },
  { v: "percentile", t: "百分点值" },
  { v: "automatic", t: "自动" },
];

function cfDataBarMinNeedsValue(t: CfDataBarMinEndpointType): boolean {
  return t === "number" || t === "percent" || t === "formula" || t === "percentile";
}

function cfDataBarMaxNeedsValue(t: CfDataBarMaxEndpointType): boolean {
  return t === "number" || t === "percent" || t === "formula" || t === "percentile";
}

function cfEndpointNeedsValue(t: CfColorScaleEndpointType): boolean {
  return t === "number" || t === "percent" || t === "formula" || t === "percentile";
}

function positionCfPalettePop(pop: HTMLElement, anchor: HTMLElement): void {
  const r = anchor.getBoundingClientRect();
  const margin = 8;
  const pw = pop.offsetWidth || 220;
  const ph = pop.offsetHeight || 280;
  let left = r.left;
  let top = r.bottom + 4;
  if (left + pw + margin > window.innerWidth) {
    left = Math.max(margin, window.innerWidth - pw - margin);
  }
  if (top + ph + margin > window.innerHeight) {
    top = Math.max(margin, r.top - ph - 4);
  }
  pop.style.left = `${Math.min(left, window.innerWidth - pw - margin)}px`;
  pop.style.top = `${Math.min(top, window.innerHeight - ph - margin)}px`;
}

/** 在锚点下方展开色盘；返回关闭函数。 */
function mountCfColorPopover(
  anchor: HTMLElement,
  currentCssHex: string,
  onPick: (cssHex: string) => void,
): () => void {
  const backdrop = document.createElement("div");
  backdrop.className = "fs-cf-color-popover-layer";
  const pop = document.createElement("div");
  pop.className = "fs-color-menu fs-color-menu--cf-dialog-pop";
  pop.setAttribute("role", "menu");

  let removed = false;
  const remove = (): void => {
    if (removed) {
      return;
    }
    removed = true;
    backdrop.remove();
    pop.remove();
    window.removeEventListener("keydown", onEsc, true);
    window.removeEventListener("scroll", reposition, true);
  };

  const onEsc = (e: KeyboardEvent): void => {
    if (e.key === "Escape") {
      e.stopImmediatePropagation();
      e.preventDefault();
      remove();
    }
  };

  const reposition = (): void => {
    positionCfPalettePop(pop, anchor);
  };

  appendRibbonColorPaletteContent(pop, {
    themeHeading: "主题颜色",
    standardHeading: "标准色",
    onPickHex: (h) => {
      onPick(h.toLowerCase());
      remove();
    },
    onMoreColors: () => {
      remove();
      void (async () => {
        const m = await showRibbonColorDialog(currentCssHex);
        if (m !== null) {
          onPick(m.toLowerCase());
        }
      })();
    },
  });

  document.body.appendChild(backdrop);
  document.body.appendChild(pop);
  requestAnimationFrame(() => {
    positionCfPalettePop(pop, anchor);
  });
  backdrop.addEventListener("mousedown", (e) => {
    if (e.target === backdrop) {
      remove();
    }
  });
  window.addEventListener("keydown", onEsc, true);
  window.addEventListener("scroll", reposition, true);
  return remove;
}

export type CfNewRuleDialogSeed =
  | { readonly kind: "default" }
  | {
      readonly kind: "highlightPreset";
      readonly highlightCommandId: string;
    };

function mapHighlightCommandToSeed(
  id: string,
): Pick<
  ConditionalFormatRule,
  "classicType" | "cellsThatContainKind" | "valueOperator" | "textOperator" | "uniqueKind"
> | null {
  switch (id) {
    case "home.style.conditional.highlightCells.greaterThan":
      return {
        classicType: "cellsThatContain",
        cellsThatContainKind: "cellValue",
        valueOperator: "greaterThan",
      };
    case "home.style.conditional.highlightCells.lessThan":
      return {
        classicType: "cellsThatContain",
        cellsThatContainKind: "cellValue",
        valueOperator: "lessThan",
      };
    case "home.style.conditional.highlightCells.between":
      return {
        classicType: "cellsThatContain",
        cellsThatContainKind: "cellValue",
        valueOperator: "between",
      };
    case "home.style.conditional.highlightCells.equalTo":
      return {
        classicType: "cellsThatContain",
        cellsThatContainKind: "cellValue",
        valueOperator: "equal",
      };
    case "home.style.conditional.highlightCells.textContains":
      return {
        classicType: "cellsThatContain",
        cellsThatContainKind: "specificText",
        textOperator: "contains",
      };
    case "home.style.conditional.highlightCells.dateOccurring":
      return { classicType: "cellsThatContain", cellsThatContainKind: "dateOccurring" };
    case "home.style.conditional.highlightCells.duplicateValues":
      return { classicType: "uniqueOrDuplicate", uniqueKind: "duplicate" };
    default:
      return null;
  }
}

/**
 * 打开「新建格式规则」对话框；确定返回规则（含 `id`），取消返回 `null`。
 * `range` 为规则作用范围（通常为当前选区）。
 */
export function showConditionalFormatNewRuleDialog(
  range: SelectionRange,
  seed: CfNewRuleDialogSeed = { kind: "default" },
): Promise<ConditionalFormatRule | null> {
  return new Promise((resolve) => {
    let settled = false;
    const finish = (v: ConditionalFormatRule | null): void => {
      if (settled) {
        return;
      }
      settled = true;
      teardown?.();
      resolve(v);
    };

    let seedPartial: ReturnType<typeof mapHighlightCommandToSeed> | null = null;
    if (seed.kind === "highlightPreset") {
      seedPartial = mapHighlightCommandToSeed(seed.highlightCommandId);
    }

    const backdrop = document.createElement("div");
    backdrop.className = "fs-cf-dialog-backdrop";
    if (ribbonIsDarkTheme()) {
      backdrop.classList.add("fs-cf-dialog-backdrop--dark");
    }

    const dlg = document.createElement("div");
    dlg.className = "fs-cf-dialog";
    dlg.setAttribute("role", "dialog");
    dlg.setAttribute("aria-modal", "true");
    dlg.setAttribute("aria-labelledby", "fs-cf-dialog-title");

    const header = document.createElement("div");
    header.className = "fs-cf-dialog__header";
    const title = document.createElement("div");
    title.className = "fs-cf-dialog__title";
    title.id = "fs-cf-dialog-title";
    title.textContent = "新建格式规则";
    const closeBtn = document.createElement("button");
    closeBtn.type = "button";
    closeBtn.className = "fs-cf-dialog__close";
    closeBtn.setAttribute("aria-label", "关闭");
    closeBtn.textContent = "\u00D7";
    closeBtn.addEventListener("click", () => finish(null));
    header.appendChild(title);
    header.appendChild(closeBtn);

    const body = document.createElement("div");
    body.className = "fs-cf-dialog__body";

    const rowStyle = document.createElement("div");
    rowStyle.className = "fs-cf-dialog__row";
    const labStyle = document.createElement("label");
    labStyle.className = "fs-cf-dialog__label";
    labStyle.textContent = "样式:";
    labStyle.setAttribute("for", "fs-cf-ui-family");
    const selUiFamily = document.createElement("select");
    selUiFamily.id = "fs-cf-ui-family";
    selUiFamily.className = "fs-cf-dialog__select";
    const uiOpts: { v: ConditionalFormatUiFamily; t: string }[] = [
      { v: "twoColorScale", t: "双色刻度" },
      { v: "threeColorScale", t: "三色刻度" },
      { v: "dataBar", t: "数据条" },
      { v: "iconSet", t: "图标集" },
      { v: "classic", t: "经典" },
    ];
    for (const o of uiOpts) {
      const op = document.createElement("option");
      op.value = o.v;
      op.textContent = o.t;
      selUiFamily.appendChild(op);
    }
    selUiFamily.value = "classic";
    rowStyle.appendChild(labStyle);
    rowStyle.appendChild(selUiFamily);

    const tipNonClassic = document.createElement("div");
    tipNonClassic.className = "fs-cf-dialog__tip";
    tipNonClassic.hidden = true;
    tipNonClassic.textContent = "当前版本暂不支持「图标集」样式。";

    /** 仅挂载当前样式对应表单（经典 / 双色刻度互斥，不并存 DOM）。 */
    const styleHost = document.createElement("div");
    styleHost.className = "fs-cf-dialog__style-host";

    const footer = document.createElement("div");
    footer.className = "fs-cf-dialog__footer";
    const cancel = document.createElement("button");
    cancel.type = "button";
    cancel.className = "fs-cf-dialog__btn fs-cf-dialog__btn--secondary";
    cancel.textContent = "取消";
    cancel.addEventListener("click", () => finish(null));
    const ok = document.createElement("button");
    ok.type = "button";
    ok.className = "fs-cf-dialog__btn fs-cf-dialog__btn--primary";
    ok.textContent = "确定";

    footer.appendChild(cancel);
    footer.appendChild(ok);

    body.appendChild(rowStyle);
    body.appendChild(tipNonClassic);
    body.appendChild(styleHost);

    dlg.appendChild(header);
    dlg.appendChild(body);
    dlg.appendChild(footer);
    backdrop.appendChild(dlg);
    document.body.appendChild(backdrop);

    const valueOpLabels: { v: CfValueOperator; t: string }[] = [
      { v: "between", t: "介于" },
      { v: "notBetween", t: "未介于" },
      { v: "equal", t: "等于" },
      { v: "notEqual", t: "不等于" },
      { v: "greaterThan", t: "大于" },
      { v: "lessThan", t: "小于" },
      { v: "greaterThanOrEqual", t: "大于或等于" },
      { v: "lessThanOrEqual", t: "小于或等于" },
    ];

    type CfNewRuleLive =
      | {
          readonly k: "classic";
          readonly dynamic: HTMLDivElement;
          readonly selClassicType: HTMLSelectElement;
          readonly selPreset: HTMLSelectElement;
          readonly preview: HTMLDivElement;
          readonly customBtn: HTMLButtonElement;
          customFillHex: string | null;
          customFgHex: string | null;
          rebuildDynamic: () => void;
          syncPreview: () => void;
        }
      | {
          readonly k: "two";
          readonly sMin: {
            readonly col: HTMLDivElement;
            readonly st: HTMLSelectElement;
            readonly inp: HTMLInputElement;
            readonly cbtn: HTMLButtonElement;
            readonly syncBar: (hex: string) => void;
          };
          readonly sMax: {
            readonly col: HTMLDivElement;
            readonly st: HTMLSelectElement;
            readonly inp: HTMLInputElement;
            readonly cbtn: HTMLButtonElement;
            readonly syncBar: (hex: string) => void;
          };
          scaleMinHex: string;
          scaleMaxHex: string;
          readonly syncScaleValUi: () => void;
        }
      | {
          readonly k: "three";
          readonly sMin: {
            readonly col: HTMLDivElement;
            readonly st: HTMLSelectElement;
            readonly inp: HTMLInputElement;
            readonly cbtn: HTMLButtonElement;
            readonly syncBar: (hex: string) => void;
          };
          readonly sMid: {
            readonly col: HTMLDivElement;
            readonly st: HTMLSelectElement;
            readonly inp: HTMLInputElement;
            readonly cbtn: HTMLButtonElement;
            readonly syncBar: (hex: string) => void;
          };
          readonly sMax: {
            readonly col: HTMLDivElement;
            readonly st: HTMLSelectElement;
            readonly inp: HTMLInputElement;
            readonly cbtn: HTMLButtonElement;
            readonly syncBar: (hex: string) => void;
          };
          scaleMinHex: string;
          scaleMidHex: string;
          scaleMaxHex: string;
          readonly syncScaleValUi: () => void;
        }
      | {
          readonly k: "dataBar";
          readonly sMin: {
            readonly col: HTMLDivElement;
            readonly st: HTMLSelectElement;
            readonly inp: HTMLInputElement;
          };
          readonly sMax: {
            readonly col: HTMLDivElement;
            readonly st: HTMLSelectElement;
            readonly inp: HTMLInputElement;
          };
          readonly selDir: HTMLSelectElement;
          readonly chkShowBarOnly: HTMLInputElement;
          readonly fillSolid: HTMLInputElement;
          readonly fillGrad: HTMLInputElement;
          readonly borderSolid: HTMLInputElement;
          readonly borderNone: HTMLInputElement;
          readonly cFillPos: HTMLButtonElement;
          readonly syncFillPos: (hex: string) => void;
          readonly cFillNeg: HTMLButtonElement;
          readonly syncFillNeg: (hex: string) => void;
          readonly cBrPos: HTMLButtonElement;
          readonly syncBrPos: (hex: string) => void;
          readonly cBrNeg: HTMLButtonElement;
          readonly syncBrNeg: (hex: string) => void;
          readonly selAxis: HTMLSelectElement;
          readonly cAxis: HTMLButtonElement;
          readonly syncAxis: (hex: string) => void;
          posFillHex: string;
          negFillHex: string;
          posBorderHex: string;
          negBorderHex: string;
          axisHex: string;
          readonly syncDataBarValueUi: () => void;
        };

    let live: CfNewRuleLive | null = null;
    let scaleCloseColorPop: (() => void) | null = null;

    const mkScaleEp = (
      hdr: string,
      opts: readonly { readonly v: CfColorScaleEndpointType; readonly t: string }[],
      def: CfColorScaleEndpointType,
    ): {
      readonly col: HTMLDivElement;
      readonly st: HTMLSelectElement;
      readonly inp: HTMLInputElement;
      readonly cbtn: HTMLButtonElement;
      readonly syncBar: (hex: string) => void;
    } => {
      const col = document.createElement("div");
      col.className = "fs-cf-dialog__scale-col";
      const head = document.createElement("div");
      head.className = "fs-cf-dialog__scale-colhead";
      head.textContent = hdr;

      const rt = document.createElement("div");
      rt.className = "fs-cf-dialog__row fs-cf-dialog__row--scale-type";
      const lt = document.createElement("span");
      lt.className = "fs-cf-dialog__label fs-cf-dialog__label--narrow";
      lt.textContent = "类型:";
      const st = document.createElement("select");
      st.className = "fs-cf-dialog__select";
      for (const o of opts) {
        const op = document.createElement("option");
        op.value = o.v;
        op.textContent = o.t;
        st.appendChild(op);
      }
      st.value = def;
      rt.appendChild(lt);
      rt.appendChild(st);

      const rv = document.createElement("div");
      rv.className = "fs-cf-dialog__row fs-cf-dialog__row--scale-val";
      const lv = document.createElement("span");
      lv.className = "fs-cf-dialog__label fs-cf-dialog__label--narrow";
      lv.textContent = "值:";
      const wrap = document.createElement("div");
      wrap.className = "fs-cf-dialog__value-combo";
      const inp = document.createElement("input");
      inp.type = "text";
      inp.className = "fs-cf-dialog__input";
      const ref = document.createElement("button");
      ref.type = "button";
      ref.className = "fs-cf-dialog__range-pick";
      ref.title = "从表格选取引用（暂不支持）";
      ref.setAttribute("aria-label", "从表格选取");
      ref.textContent = "\u25A4";
      wrap.appendChild(inp);
      wrap.appendChild(ref);

      rv.appendChild(lv);
      rv.appendChild(wrap);

      const rc = document.createElement("div");
      rc.className = "fs-cf-dialog__row fs-cf-dialog__row--scale-color";
      const lc = document.createElement("span");
      lc.className = "fs-cf-dialog__label fs-cf-dialog__label--narrow";
      lc.textContent = "颜色:";
      const cbtn = document.createElement("button");
      cbtn.type = "button";
      cbtn.className = "fs-cf-dialog__color-swatch-btn";
      const cbar = document.createElement("span");
      cbar.className = "fs-cf-dialog__color-swatch-btn__bar";
      const cchev = document.createElement("span");
      cchev.className = "fs-cf-dialog__color-swatch-btn__chev";
      cchev.textContent = "\u25BE";
      cbtn.appendChild(cbar);
      cbtn.appendChild(cchev);

      rc.appendChild(lc);
      rc.appendChild(cbtn);

      col.appendChild(head);
      col.appendChild(rt);
      col.appendChild(rv);
      col.appendChild(rc);

      const syncBar = (hex: string): void => {
        cbar.style.backgroundColor = hex;
      };

      return { col, st, inp, cbtn, syncBar };
    };

    const mkDataBarBoundCol = (
      hdr: string,
      opts: readonly { readonly v: string; readonly t: string }[],
      def: string,
    ): {
      readonly col: HTMLDivElement;
      readonly st: HTMLSelectElement;
      readonly inp: HTMLInputElement;
    } => {
      const col = document.createElement("div");
      col.className = "fs-cf-dialog__scale-col";
      const head = document.createElement("div");
      head.className = "fs-cf-dialog__scale-colhead";
      head.textContent = hdr;
      const rt = document.createElement("div");
      rt.className = "fs-cf-dialog__row fs-cf-dialog__row--scale-type";
      const lt = document.createElement("span");
      lt.className = "fs-cf-dialog__label fs-cf-dialog__label--narrow";
      lt.textContent = "类型:";
      const st = document.createElement("select");
      st.className = "fs-cf-dialog__select";
      for (const o of opts) {
        const op = document.createElement("option");
        op.value = o.v;
        op.textContent = o.t;
        st.appendChild(op);
      }
      st.value = def;
      rt.appendChild(lt);
      rt.appendChild(st);
      const rv = document.createElement("div");
      rv.className = "fs-cf-dialog__row fs-cf-dialog__row--scale-val";
      const lv = document.createElement("span");
      lv.className = "fs-cf-dialog__label fs-cf-dialog__label--narrow";
      lv.textContent = "值:";
      const wrap = document.createElement("div");
      wrap.className = "fs-cf-dialog__value-combo";
      const inp = document.createElement("input");
      inp.type = "text";
      inp.className = "fs-cf-dialog__input";
      const ref = document.createElement("button");
      ref.type = "button";
      ref.className = "fs-cf-dialog__range-pick";
      ref.title = "从表格选取引用（暂不支持）";
      ref.setAttribute("aria-label", "从表格选取");
      ref.textContent = "\u25A4";
      wrap.appendChild(inp);
      wrap.appendChild(ref);
      rv.appendChild(lv);
      rv.appendChild(wrap);
      col.appendChild(head);
      col.appendChild(rt);
      col.appendChild(rv);
      return { col, st, inp };
    };

    const disposeStyleForm = (): void => {
      scaleCloseColorPop?.();
      scaleCloseColorPop = null;
      styleHost.replaceChildren();
      live = null;
    };

    const rebuildClassicDynamic = (
      dynamic: HTMLDivElement,
      selClassicType: HTMLSelectElement,
      onRefresh: () => void,
    ): void => {
      dynamic.textContent = "";
      const ct = selClassicType.value as ConditionalFormatClassicRuleType;
      if (ct === "formula") {
        const p = document.createElement("p");
        p.className = "fs-cf-dialog__warn";
        p.textContent = "当前版本暂不支持「使用公式确定要设置格式的单元格」。";
        dynamic.appendChild(p);
        return;
      }
      if (ct === "cellsThatContain") {
        const row1 = document.createElement("div");
        row1.className = "fs-cf-dialog__row";
        const s1 = document.createElement("select");
        s1.className = "fs-cf-dialog__select";
        const kinds: { v: CfCellsThatContainKind; t: string }[] = [
          { v: "cellValue", t: "单元格值" },
          { v: "specificText", t: "特定文本" },
          { v: "dateOccurring", t: "发生日期" },
          { v: "blanks", t: "空值" },
          { v: "noBlanks", t: "无空值" },
          { v: "errors", t: "错误" },
          { v: "noErrors", t: "无错误" },
        ];
        for (const k of kinds) {
          const op = document.createElement("option");
          op.value = k.v;
          op.textContent = k.t;
          s1.appendChild(op);
        }
        s1.value = seedPartial?.cellsThatContainKind ?? "cellValue";

        const s2 = document.createElement("select");
        s2.className = "fs-cf-dialog__select";
        for (const o of valueOpLabels) {
          const op = document.createElement("option");
          op.value = o.v;
          op.textContent = o.t;
          s2.appendChild(op);
        }
        s2.value = (seedPartial?.valueOperator ?? "greaterThan") as CfValueOperator;

        const inp1 = document.createElement("input");
        inp1.type = "text";
        inp1.className = "fs-cf-dialog__input";
        inp1.placeholder = "值或引用";

        const inp2 = document.createElement("input");
        inp2.type = "text";
        inp2.className = "fs-cf-dialog__input";
        inp2.placeholder = "第二个值（介于）";

        const st = document.createElement("select");
        st.className = "fs-cf-dialog__select";
        for (const o of [
          { v: "contains" as const, t: "包含" },
          { v: "notContains" as const, t: "不包含" },
          { v: "beginsWith" as const, t: "开始于" },
          { v: "endsWith" as const, t: "结束于" },
        ]) {
          const op = document.createElement("option");
          op.value = o.v;
          op.textContent = o.t;
          st.appendChild(op);
        }
        st.value = (seedPartial?.textOperator ?? "contains") as CfTextOperator;

        const sd = document.createElement("select");
        sd.className = "fs-cf-dialog__select";
        for (const d of [
          { v: "today" as const, t: "今天" },
          { v: "yesterday" as const, t: "昨天" },
          { v: "tomorrow" as const, t: "明天" },
          { v: "thisWeek" as const, t: "本周" },
          { v: "lastWeek" as const, t: "上周" },
          { v: "nextWeek" as const, t: "下周" },
          { v: "thisMonth" as const, t: "本月" },
          { v: "lastMonth" as const, t: "上月" },
          { v: "nextMonth" as const, t: "下月" },
        ]) {
          const op = document.createElement("option");
          op.value = d.v;
          op.textContent = d.t;
          sd.appendChild(op);
        }

        const applyKindLayout = (): void => {
          while (row1.firstChild) {
            row1.removeChild(row1.firstChild);
          }
          const k = s1.value as CfCellsThatContainKind;
          row1.appendChild(s1);
          if (k === "cellValue") {
            row1.appendChild(s2);
            row1.appendChild(inp1);
            const need2 = s2.value === "between" || s2.value === "notBetween";
            if (need2) {
              row1.appendChild(inp2);
            }
          } else if (k === "specificText") {
            row1.appendChild(st);
            row1.appendChild(inp1);
          } else if (k === "dateOccurring") {
            row1.appendChild(sd);
          }
        };

        s1.addEventListener("change", () => {
          applyKindLayout();
          onRefresh();
        });
        s2.addEventListener("change", () => {
          applyKindLayout();
          onRefresh();
        });

        applyKindLayout();
        dynamic.appendChild(row1);

        const bag = dynamic as unknown as {
          _ctKind: HTMLSelectElement;
          _valOp: HTMLSelectElement;
          _inp1: HTMLInputElement;
          _inp2: HTMLInputElement;
          _textOp: HTMLSelectElement;
          _dateSel: HTMLSelectElement;
        };
        bag._ctKind = s1;
        bag._valOp = s2;
        bag._inp1 = inp1;
        bag._inp2 = inp2;
        bag._textOp = st;
        bag._dateSel = sd;
        return;
      }
      if (ct === "topBottomRanked") {
        const row = document.createElement("div");
        row.className = "fs-cf-dialog__row";
        const sk = document.createElement("select");
        sk.className = "fs-cf-dialog__select";
        for (const o of [
          { v: "top" as const, t: "最大" },
          { v: "bottom" as const, t: "最小" },
          { v: "topPercent" as const, t: "前 10%" },
          { v: "bottomPercent" as const, t: "后 10%" },
        ]) {
          const op = document.createElement("option");
          op.value = o.v;
          op.textContent = o.t;
          sk.appendChild(op);
        }
        const num = document.createElement("input");
        num.type = "number";
        num.className = "fs-cf-dialog__input fs-cf-dialog__input--narrow";
        num.min = "1";
        num.max = "1000";
        num.value = "10";
        row.appendChild(sk);
        row.appendChild(num);
        dynamic.appendChild(row);
        (dynamic as unknown as { _tbKind?: HTMLSelectElement })._tbKind = sk;
        (dynamic as unknown as { _tbN?: HTMLInputElement })._tbN = num;
        return;
      }
      if (ct === "aboveBelowAverage") {
        const sk = document.createElement("select");
        sk.className = "fs-cf-dialog__select fs-cf-dialog__select--wide";
        for (const o of [
          { v: "above" as const, t: "高于平均值" },
          { v: "below" as const, t: "低于平均值" },
          { v: "equalOrAbove" as const, t: "等于或高于平均值" },
          { v: "equalOrBelow" as const, t: "等于或低于平均值" },
        ]) {
          const op = document.createElement("option");
          op.value = o.v;
          op.textContent = o.t;
          sk.appendChild(op);
        }
        dynamic.appendChild(sk);
        (dynamic as unknown as { _avgKind?: HTMLSelectElement })._avgKind = sk;
        return;
      }
      if (ct === "uniqueOrDuplicate") {
        const sk = document.createElement("select");
        sk.className = "fs-cf-dialog__select fs-cf-dialog__select--wide";
        for (const o of [
          { v: "duplicate" as const, t: "重复" },
          { v: "unique" as const, t: "唯一" },
        ]) {
          const op = document.createElement("option");
          op.value = o.v;
          op.textContent = o.t;
          sk.appendChild(op);
        }
        sk.value = (seedPartial?.uniqueKind ?? "duplicate") as CfUniqueKind;
        dynamic.appendChild(sk);
        (dynamic as unknown as { _uniqKind?: HTMLSelectElement })._uniqKind = sk;
      }
    };

    const readCellsThatContainRule = (): Omit<
      ConditionalFormatRule,
      "id" | "range" | "uiFamily" | "formatPreset" | "customFormat"
    > => {
      const dyn = live?.k === "classic" ? live.dynamic : null;
      const bag = dyn as unknown as {
        _ctKind?: HTMLSelectElement;
        _valOp?: HTMLSelectElement;
        _inp1?: HTMLInputElement;
        _inp2?: HTMLInputElement;
        _textOp?: HTMLSelectElement;
        _dateSel?: HTMLSelectElement;
      };
      const kind = (bag._ctKind?.value ?? "cellValue") as CfCellsThatContainKind;
      const base: Omit<
        ConditionalFormatRule,
        "id" | "range" | "uiFamily" | "formatPreset" | "customFormat"
      > = {
        classicType: "cellsThatContain",
        cellsThatContainKind: kind,
      };
      if (kind === "cellValue") {
        return {
          ...base,
          valueOperator: (bag._valOp?.value ?? "greaterThan") as CfValueOperator,
          value1: bag._inp1?.value ?? "",
          value2: bag._inp2?.value ?? "",
        };
      }
      if (kind === "specificText") {
        return {
          ...base,
          textOperator: (bag._textOp?.value ?? "contains") as CfTextOperator,
          value1: bag._inp1?.value ?? "",
        };
      }
      if (kind === "dateOccurring") {
        return {
          ...base,
          dateOccurring: (bag._dateSel?.value ?? "today") as CfDateOccurring,
        };
      }
      return base;
    };

    const readRuleFromForm = (): ConditionalFormatRule | null => {
      const ui = selUiFamily.value as ConditionalFormatUiFamily;
      if (ui === "twoColorScale") {
        if (live?.k !== "two") {
          return null;
        }
        return {
          id: newCfRuleId(),
          range,
          uiFamily: "twoColorScale",
          classicType: "colorScale",
          formatPreset: "none",
          cfTwoColorMin: {
            type: live.sMin.st.value as CfColorScaleEndpointType,
            value: live.sMin.inp.value.trim(),
            colorArgb: cssHexToFillArgb(live.scaleMinHex),
          },
          cfTwoColorMax: {
            type: live.sMax.st.value as CfColorScaleEndpointType,
            value: live.sMax.inp.value.trim(),
            colorArgb: cssHexToFillArgb(live.scaleMaxHex),
          },
        };
      }
      if (ui === "threeColorScale") {
        if (live?.k !== "three") {
          return null;
        }
        return {
          id: newCfRuleId(),
          range,
          uiFamily: "threeColorScale",
          classicType: "colorScale",
          formatPreset: "none",
          cfThreeColorMin: {
            type: live.sMin.st.value as CfColorScaleEndpointType,
            value: live.sMin.inp.value.trim(),
            colorArgb: cssHexToFillArgb(live.scaleMinHex),
          },
          cfThreeColorMid: {
            type: live.sMid.st.value as CfColorScaleEndpointType,
            value: live.sMid.inp.value.trim(),
            colorArgb: cssHexToFillArgb(live.scaleMidHex),
          },
          cfThreeColorMax: {
            type: live.sMax.st.value as CfColorScaleEndpointType,
            value: live.sMax.inp.value.trim(),
            colorArgb: cssHexToFillArgb(live.scaleMaxHex),
          },
        };
      }
      if (ui === "dataBar") {
        if (live?.k !== "dataBar") {
          return null;
        }
        const d = live;
        const fillKind: CfDataBarFillKind = d.fillSolid.checked ? "solid" : "gradient";
        const borderKind: CfDataBarBorderKind = d.borderNone.checked ? "none" : "solid";
        const base: ConditionalFormatRule = {
          id: newCfRuleId(),
          range,
          uiFamily: "dataBar",
          classicType: "dataBar",
          formatPreset: "none",
          cfDataBarMin: {
            type: d.sMin.st.value as CfDataBarMinEndpointType,
            value: d.sMin.inp.value.trim(),
          },
          cfDataBarMax: {
            type: d.sMax.st.value as CfDataBarMaxEndpointType,
            value: d.sMax.inp.value.trim(),
          },
          cfDataBarDirection: d.selDir.value as CfDataBarDirection,
          cfDataBarShowBarOnly: d.chkShowBarOnly.checked,
          cfDataBarFillKind: fillKind,
          cfDataBarPositiveFillArgb: cssHexToFillArgb(d.posFillHex),
          cfDataBarNegativeFillArgb: cssHexToFillArgb(d.negFillHex),
          cfDataBarBorderKind: borderKind,
          cfDataBarAxisPosition: d.selAxis.value as CfDataBarAxisPosition,
          cfDataBarAxisColorArgb: cssHexToFillArgb(d.axisHex),
        };
        if (borderKind === "solid") {
          return {
            ...base,
            cfDataBarPositiveBorderArgb: cssHexToFillArgb(d.posBorderHex),
            cfDataBarNegativeBorderArgb: cssHexToFillArgb(d.negBorderHex),
          };
        }
        return base;
      }
      if (ui !== "classic" || live?.k !== "classic") {
        return null;
      }
      const c = live;
      const preset = c.selPreset.value as CfFormatPresetId;
      let customFormat: ConditionalFormatRule["customFormat"];
      if (preset === "custom") {
        if (c.customFillHex === null && c.customFgHex === null) {
          customFormat = { fillArgb: "FFFFC7CE", fgArgb: "FF9C0006" };
        } else {
          customFormat = {
            ...(c.customFillHex !== null ? { fillArgb: cssHexToFillArgb(c.customFillHex) } : {}),
            ...(c.customFgHex !== null ? { fgArgb: cssHexToFillArgb(c.customFgHex) } : {}),
          } as CellStylePatch;
        }
      }
      const ct = c.selClassicType.value as ConditionalFormatClassicRuleType;
      if (ct === "formula") {
        return null;
      }
      const common = {
        id: newCfRuleId(),
        range,
        uiFamily: ui,
        formatPreset: preset,
        ...(preset === "custom" ? { customFormat: customFormat! } : {}),
      } as const;

      if (ct === "cellsThatContain") {
        return { ...common, ...readCellsThatContainRule() } as ConditionalFormatRule;
      }
      if (ct === "topBottomRanked") {
        const bag = c.dynamic as unknown as {
          _tbKind?: HTMLSelectElement;
          _tbN?: HTMLInputElement;
        };
        const n = Math.max(1, Math.min(1000, Number(bag._tbN?.value ?? 10) || 10));
        return {
          ...common,
          classicType: "topBottomRanked",
          topBottomKind: (bag._tbKind?.value ?? "top") as CfTopBottomKind,
          topBottomN: n,
        };
      }
      if (ct === "aboveBelowAverage") {
        const bag = c.dynamic as unknown as { _avgKind?: HTMLSelectElement };
        return {
          ...common,
          classicType: "aboveBelowAverage",
          averageKind: (bag._avgKind?.value ?? "above") as CfAverageKind,
        };
      }
      if (ct === "uniqueOrDuplicate") {
        const bag = c.dynamic as unknown as { _uniqKind?: HTMLSelectElement };
        return {
          ...common,
          classicType: "uniqueOrDuplicate",
          uniqueKind: (bag._uniqKind?.value ?? "duplicate") as CfUniqueKind,
        };
      }
      return null;
    };

    const validate = (): boolean => {
      const ui = selUiFamily.value as ConditionalFormatUiFamily;
      if (ui === "twoColorScale") {
        if (live?.k !== "two") {
          return false;
        }
        const t = live;
        const checkEp = (ep: CfColorScaleEndpointType, v: string): boolean => {
          if (!cfEndpointNeedsValue(ep)) {
            return true;
          }
          const s = v.trim();
          if (s === "") {
            return false;
          }
          if (ep === "formula") {
            return true;
          }
          const n = Number(s.replace(/^=/, ""));
          return Number.isFinite(n);
        };
        return (
          checkEp(t.sMin.st.value as CfColorScaleEndpointType, t.sMin.inp.value) &&
          checkEp(t.sMax.st.value as CfColorScaleEndpointType, t.sMax.inp.value)
        );
      }
      if (ui === "threeColorScale") {
        if (live?.k !== "three") {
          return false;
        }
        const t = live;
        const checkEp = (ep: CfColorScaleEndpointType, v: string): boolean => {
          if (!cfEndpointNeedsValue(ep)) {
            return true;
          }
          const s = v.trim();
          if (s === "") {
            return false;
          }
          if (ep === "formula") {
            return true;
          }
          const n = Number(s.replace(/^=/, ""));
          return Number.isFinite(n);
        };
        return (
          checkEp(t.sMin.st.value as CfColorScaleEndpointType, t.sMin.inp.value) &&
          checkEp(t.sMid.st.value as CfColorScaleEndpointType, t.sMid.inp.value) &&
          checkEp(t.sMax.st.value as CfColorScaleEndpointType, t.sMax.inp.value)
        );
      }
      if (ui === "dataBar") {
        if (live?.k !== "dataBar") {
          return false;
        }
        const t = live;
        const checkMin = (ep: CfDataBarMinEndpointType, v: string): boolean => {
          if (!cfDataBarMinNeedsValue(ep)) {
            return true;
          }
          const s = v.trim();
          if (s === "") {
            return false;
          }
          if (ep === "formula") {
            return true;
          }
          const n = Number(s.replace(/^=/, ""));
          return Number.isFinite(n);
        };
        const checkMax = (ep: CfDataBarMaxEndpointType, v: string): boolean => {
          if (!cfDataBarMaxNeedsValue(ep)) {
            return true;
          }
          const s = v.trim();
          if (s === "") {
            return false;
          }
          if (ep === "formula") {
            return true;
          }
          const n = Number(s.replace(/^=/, ""));
          return Number.isFinite(n);
        };
        return (
          checkMin(t.sMin.st.value as CfDataBarMinEndpointType, t.sMin.inp.value) &&
          checkMax(t.sMax.st.value as CfDataBarMaxEndpointType, t.sMax.inp.value)
        );
      }
      if (ui !== "classic" || live?.k !== "classic") {
        return false;
      }
      const ct = live.selClassicType.value as ConditionalFormatClassicRuleType;
      if (ct === "formula") {
        return false;
      }
      if (ct === "cellsThatContain") {
        const r = readCellsThatContainRule();
        if (r.cellsThatContainKind === "cellValue") {
          const op = r.valueOperator;
          if (op === "between" || op === "notBetween") {
            return (r.value1?.trim() ?? "") !== "" && (r.value2?.trim() ?? "") !== "";
          }
          return (r.value1?.trim() ?? "") !== "";
        }
        if (r.cellsThatContainKind === "specificText") {
          return (r.value1?.trim() ?? "") !== "";
        }
        return true;
      }
      return true;
    };

    const refreshOk = (): void => {
      ok.disabled = !validate();
    };

    const paintStyleForm = (): CfNewRuleLive | null => {
      disposeStyleForm();
      const ui = selUiFamily.value as ConditionalFormatUiFamily;
      if (
        ui !== "classic" &&
        ui !== "twoColorScale" &&
        ui !== "threeColorScale" &&
        ui !== "dataBar"
      ) {
        return null;
      }
      if (ui === "classic") {
        const classicPanel = document.createElement("div");
        classicPanel.className = "fs-cf-dialog__panel";
        const rowRuleType = document.createElement("div");
        rowRuleType.className = "fs-cf-dialog__row";
        const selClassicType = document.createElement("select");
        selClassicType.className = "fs-cf-dialog__select fs-cf-dialog__select--wide";
        const classicTypeOpts: { v: ConditionalFormatClassicRuleType; t: string }[] = [
          { v: "cellsThatContain", t: "只为包含以下内容的单元格设置格式" },
          { v: "topBottomRanked", t: "仅对排名靠前或靠后的数值设置格式" },
          { v: "aboveBelowAverage", t: "仅对高于或低于平均值的数值设置格式" },
          { v: "uniqueOrDuplicate", t: "仅对唯一值或重复值设置格式" },
          { v: "formula", t: "使用公式确定要设置格式的单元格" },
        ];
        for (const o of classicTypeOpts) {
          const op = document.createElement("option");
          op.value = o.v;
          op.textContent = o.t;
          selClassicType.appendChild(op);
        }
        rowRuleType.appendChild(selClassicType);

        const dynamic = document.createElement("div");
        dynamic.className = "fs-cf-dialog__dynamic";

        const rowFmt = document.createElement("div");
        rowFmt.className = "fs-cf-dialog__row fs-cf-dialog__row--format";
        const labFmt = document.createElement("label");
        labFmt.className = "fs-cf-dialog__label";
        labFmt.textContent = "设置格式:";
        labFmt.setAttribute("for", "fs-cf-fmt-preset");
        const selPreset = document.createElement("select");
        selPreset.id = "fs-cf-fmt-preset";
        selPreset.className = "fs-cf-dialog__select";
        const presets: { v: CfFormatPresetId; t: string }[] = [
          { v: "lightRedFillDarkRedText", t: "浅红填充色深红色文本" },
          { v: "yellowFillDarkYellowText", t: "黄填充色深黄色文本" },
          { v: "greenFillDarkGreenText", t: "绿填充色深绿色文本" },
          { v: "lightRedFill", t: "浅红色填充" },
          { v: "redText", t: "红色文本" },
          { v: "redBorder", t: "红色边框" },
          { v: "custom", t: "自定义格式..." },
        ];
        for (const p of presets) {
          const op = document.createElement("option");
          op.value = p.v;
          op.textContent = p.t;
          selPreset.appendChild(op);
        }

        const preview = document.createElement("div");
        preview.className = "fs-cf-dialog__preview";
        preview.textContent = PREVIEW_TEXT;

        const customBtn = document.createElement("button");
        customBtn.type = "button";
        customBtn.className = "fs-cf-dialog__btn-mini";
        customBtn.textContent = "自定义颜色…";
        customBtn.hidden = true;

        const classicLive: CfNewRuleLive = {
          k: "classic",
          dynamic,
          selClassicType,
          selPreset,
          preview,
          customBtn,
          customFillHex: null,
          customFgHex: null,
          rebuildDynamic: () => {
            rebuildClassicDynamic(dynamic, selClassicType, refreshOk);
          },
          syncPreview: (): void => {
            return;
          },
        };

        classicLive.syncPreview = (): void => {
          const pr = selPreset.value as CfFormatPresetId;
          if (pr === "custom") {
            const patch = {
              ...(classicLive.customFillHex !== null
                ? { fillArgb: cssHexToFillArgb(classicLive.customFillHex) }
                : {}),
              ...(classicLive.customFgHex !== null
                ? { fgArgb: cssHexToFillArgb(classicLive.customFgHex) }
                : {}),
            } as CellStylePatch;
            const o = cfFormatPresetToOverlay("custom", patch);
            preview.style.background =
              o.fillArgb !== undefined ? `#${o.fillArgb.slice(2)}` : "#ffffff";
            preview.style.color = o.fgArgb !== undefined ? `#${o.fgArgb.slice(2)}` : "#000000";
          } else {
            const o = cfFormatPresetToOverlay(pr);
            preview.style.background =
              o.fillArgb !== undefined ? `#${o.fillArgb.slice(2)}` : "transparent";
            preview.style.color = o.fgArgb !== undefined ? `#${o.fgArgb.slice(2)}` : "#000000";
          }
        };

        customBtn.addEventListener("click", () => {
          void (async () => {
            const fill = await showRibbonColorDialog(classicLive.customFillHex ?? "#ffc7ce");
            if (fill !== null) {
              classicLive.customFillHex = fill;
            }
            const fg = await showRibbonColorDialog(classicLive.customFgHex ?? "#9c0006");
            if (fg !== null) {
              classicLive.customFgHex = fg;
            }
            classicLive.syncPreview();
          })();
        });
        selPreset.addEventListener("change", () => {
          customBtn.hidden = selPreset.value !== "custom";
          classicLive.syncPreview();
        });

        rowFmt.appendChild(labFmt);
        rowFmt.appendChild(selPreset);
        rowFmt.appendChild(customBtn);
        rowFmt.appendChild(preview);

        selClassicType.addEventListener("change", () => {
          classicLive.rebuildDynamic();
          refreshOk();
        });
        dynamic.addEventListener("input", refreshOk);

        classicPanel.appendChild(rowRuleType);
        classicPanel.appendChild(dynamic);
        styleHost.appendChild(classicPanel);
        styleHost.appendChild(rowFmt);

        live = classicLive;
        classicLive.rebuildDynamic();
        classicLive.syncPreview();
        return classicLive;
      }
      if (ui === "twoColorScale") {
        const scalePanel = document.createElement("div");
        scalePanel.className = "fs-cf-dialog__scale-panel";
        const scaleSep = document.createElement("div");
        scaleSep.className = "fs-cf-dialog__scale-sep";
        const scaleCols = document.createElement("div");
        scaleCols.className = "fs-cf-dialog__scale-columns";

        const sMin = mkScaleEp("最小值", CF_MIN_TYPE_OPTS, "lowest");
        const sMax = mkScaleEp("最大值", CF_MAX_TYPE_OPTS, "highest");
        const twoLive: CfNewRuleLive = {
          k: "two",
          sMin,
          sMax,
          scaleMinHex: "#ff8c00",
          scaleMaxHex: "#fff2cc",
          syncScaleValUi: (): void => {
            const tmin = sMin.st.value as CfColorScaleEndpointType;
            const tmax = sMax.st.value as CfColorScaleEndpointType;
            if (!cfEndpointNeedsValue(tmin)) {
              sMin.inp.disabled = true;
              sMin.inp.placeholder = tmin === "lowest" ? "（最低值）" : "";
              sMin.inp.value = "";
            } else {
              sMin.inp.disabled = false;
              sMin.inp.placeholder = "";
            }
            if (!cfEndpointNeedsValue(tmax)) {
              sMax.inp.disabled = true;
              sMax.inp.placeholder = tmax === "highest" ? "（最高值）" : "";
              sMax.inp.value = "";
            } else {
              sMax.inp.disabled = false;
              sMax.inp.placeholder = "";
            }
          },
        };
        sMin.syncBar(twoLive.scaleMinHex);
        sMax.syncBar(twoLive.scaleMaxHex);

        sMin.st.addEventListener("change", () => {
          twoLive.syncScaleValUi();
          refreshOk();
        });
        sMax.st.addEventListener("change", () => {
          twoLive.syncScaleValUi();
          refreshOk();
        });
        sMin.inp.addEventListener("input", refreshOk);
        sMax.inp.addEventListener("input", refreshOk);
        sMin.cbtn.addEventListener("click", () => {
          scaleCloseColorPop?.();
          scaleCloseColorPop = mountCfColorPopover(sMin.cbtn, twoLive.scaleMinHex, (h) => {
            twoLive.scaleMinHex = h;
            sMin.syncBar(twoLive.scaleMinHex);
            refreshOk();
          });
        });
        sMax.cbtn.addEventListener("click", () => {
          scaleCloseColorPop?.();
          scaleCloseColorPop = mountCfColorPopover(sMax.cbtn, twoLive.scaleMaxHex, (h) => {
            twoLive.scaleMaxHex = h;
            sMax.syncBar(twoLive.scaleMaxHex);
            refreshOk();
          });
        });

        scalePanel.appendChild(scaleSep);
        scaleCols.appendChild(sMin.col);
        scaleCols.appendChild(sMax.col);
        scalePanel.appendChild(scaleCols);
        styleHost.appendChild(scalePanel);

        live = twoLive;
        twoLive.syncScaleValUi();
        return twoLive;
      }
      if (ui === "threeColorScale") {
        const scalePanel = document.createElement("div");
        scalePanel.className = "fs-cf-dialog__scale-panel";
        const scaleSep = document.createElement("div");
        scaleSep.className = "fs-cf-dialog__scale-sep";
        const scaleCols = document.createElement("div");
        scaleCols.className = "fs-cf-dialog__scale-columns fs-cf-dialog__scale-columns--three";

        const sMin = mkScaleEp("最小值", CF_MIN_TYPE_OPTS, "lowest");
        const sMid = mkScaleEp("中间值", CF_MID_TYPE_OPTS, "percentile");
        const sMax = mkScaleEp("最大值", CF_MAX_TYPE_OPTS, "highest");
        sMid.inp.value = "50";

        const threeLive: CfNewRuleLive = {
          k: "three",
          sMin,
          sMid,
          sMax,
          scaleMinHex: "#ff8c00",
          scaleMidHex: "#fff9c4",
          scaleMaxHex: "#fffde7",
          syncScaleValUi: (): void => {
            const tmin = sMin.st.value as CfColorScaleEndpointType;
            const tmid = sMid.st.value as CfColorScaleEndpointType;
            const tmax = sMax.st.value as CfColorScaleEndpointType;
            if (!cfEndpointNeedsValue(tmin)) {
              sMin.inp.disabled = true;
              sMin.inp.placeholder = tmin === "lowest" ? "（最低值）" : "";
              sMin.inp.value = "";
            } else {
              sMin.inp.disabled = false;
              sMin.inp.placeholder = "";
            }
            if (!cfEndpointNeedsValue(tmid)) {
              sMid.inp.disabled = true;
              sMid.inp.placeholder = "";
              sMid.inp.value = "";
            } else {
              sMid.inp.disabled = false;
              sMid.inp.placeholder = "";
            }
            if (!cfEndpointNeedsValue(tmax)) {
              sMax.inp.disabled = true;
              sMax.inp.placeholder = tmax === "highest" ? "（最高值）" : "";
              sMax.inp.value = "";
            } else {
              sMax.inp.disabled = false;
              sMax.inp.placeholder = "";
            }
          },
        };
        sMin.syncBar(threeLive.scaleMinHex);
        sMid.syncBar(threeLive.scaleMidHex);
        sMax.syncBar(threeLive.scaleMaxHex);

        const wireThree = (): void => {
          threeLive.syncScaleValUi();
          refreshOk();
        };
        sMin.st.addEventListener("change", wireThree);
        sMid.st.addEventListener("change", wireThree);
        sMax.st.addEventListener("change", wireThree);
        sMin.inp.addEventListener("input", refreshOk);
        sMid.inp.addEventListener("input", refreshOk);
        sMax.inp.addEventListener("input", refreshOk);
        sMin.cbtn.addEventListener("click", () => {
          scaleCloseColorPop?.();
          scaleCloseColorPop = mountCfColorPopover(sMin.cbtn, threeLive.scaleMinHex, (h) => {
            threeLive.scaleMinHex = h;
            sMin.syncBar(threeLive.scaleMinHex);
            refreshOk();
          });
        });
        sMid.cbtn.addEventListener("click", () => {
          scaleCloseColorPop?.();
          scaleCloseColorPop = mountCfColorPopover(sMid.cbtn, threeLive.scaleMidHex, (h) => {
            threeLive.scaleMidHex = h;
            sMid.syncBar(threeLive.scaleMidHex);
            refreshOk();
          });
        });
        sMax.cbtn.addEventListener("click", () => {
          scaleCloseColorPop?.();
          scaleCloseColorPop = mountCfColorPopover(sMax.cbtn, threeLive.scaleMaxHex, (h) => {
            threeLive.scaleMaxHex = h;
            sMax.syncBar(threeLive.scaleMaxHex);
            refreshOk();
          });
        });

        scalePanel.appendChild(scaleSep);
        scaleCols.appendChild(sMin.col);
        scaleCols.appendChild(sMid.col);
        scaleCols.appendChild(sMax.col);
        scalePanel.appendChild(scaleCols);
        styleHost.appendChild(scalePanel);

        live = threeLive;
        threeLive.syncScaleValUi();
        return threeLive;
      }
      if (ui === "dataBar") {
        const wrap = document.createElement("div");
        wrap.className = "fs-cf-dialog__data-bar-panel";

        const scaleSep0 = document.createElement("div");
        scaleSep0.className = "fs-cf-dialog__scale-sep";
        const scaleCols = document.createElement("div");
        scaleCols.className = "fs-cf-dialog__scale-columns";

        const sMin = mkDataBarBoundCol("最小值", CF_DATA_BAR_MIN_OPTS, "automatic");
        const sMax = mkDataBarBoundCol("最大值", CF_DATA_BAR_MAX_OPTS, "automatic");

        const rowDir = document.createElement("div");
        rowDir.className = "fs-cf-dialog__row fs-cf-dialog__row--data-bar-dir";
        const labDir = document.createElement("span");
        labDir.className = "fs-cf-dialog__label fs-cf-dialog__label--narrow";
        labDir.textContent = "方向:";
        const selDir = document.createElement("select");
        selDir.className = "fs-cf-dialog__select";
        for (const o of [
          { v: "context" as const, t: "上下文" },
          { v: "leftToRight" as const, t: "从左到右" },
          { v: "rightToLeft" as const, t: "从右到左" },
        ]) {
          const op = document.createElement("option");
          op.value = o.v;
          op.textContent = o.t;
          selDir.appendChild(op);
        }
        selDir.value = "leftToRight";
        const chkShowBarOnly = document.createElement("input");
        chkShowBarOnly.type = "checkbox";
        chkShowBarOnly.id = "fs-cf-dbar-baronly";
        const lblBarOnly = document.createElement("label");
        lblBarOnly.className = "fs-cf-dialog__lbl-inline";
        lblBarOnly.setAttribute("for", "fs-cf-dbar-baronly");
        lblBarOnly.textContent = "仅显示数据栏";
        rowDir.appendChild(labDir);
        rowDir.appendChild(selDir);
        rowDir.appendChild(chkShowBarOnly);
        rowDir.appendChild(lblBarOnly);

        const sep1 = document.createElement("div");
        sep1.className = "fs-cf-dialog__scale-sep";

        const secFill = document.createElement("div");
        secFill.className = "fs-cf-dialog__cf-subsec";
        const headFill = document.createElement("div");
        headFill.className = "fs-cf-dialog__cf-subsec-title";
        headFill.textContent = "填充";
        const rowFill = document.createElement("div");
        rowFill.className = "fs-cf-dialog__row fs-cf-dialog__row--cf-radios";
        const fillLab = document.createElement("span");
        fillLab.className = "fs-cf-dialog__label fs-cf-dialog__label--narrow";
        fillLab.textContent = "";
        const fillSolid = document.createElement("input");
        fillSolid.type = "radio";
        fillSolid.name = "fsCfDbarFillKind";
        fillSolid.value = "solid";
        fillSolid.checked = true;
        const lblFs = document.createElement("label");
        lblFs.className = "fs-cf-dialog__inline-radio";
        lblFs.appendChild(fillSolid);
        lblFs.appendChild(document.createTextNode(" 实心填充 "));
        const fillGrad = document.createElement("input");
        fillGrad.type = "radio";
        fillGrad.name = "fsCfDbarFillKind";
        fillGrad.value = "gradient";
        const lblFg = document.createElement("label");
        lblFg.className = "fs-cf-dialog__inline-radio";
        lblFg.appendChild(fillGrad);
        lblFg.appendChild(document.createTextNode(" 渐变填充 "));
        rowFill.appendChild(fillLab);
        rowFill.appendChild(lblFs);
        rowFill.appendChild(lblFg);

        const mkMiniSwatch = (): {
          readonly btn: HTMLButtonElement;
          readonly sync: (hex: string) => void;
        } => {
          const cbtn = document.createElement("button");
          cbtn.type = "button";
          cbtn.className = "fs-cf-dialog__color-swatch-btn";
          const cbar = document.createElement("span");
          cbar.className = "fs-cf-dialog__color-swatch-btn__bar";
          const cchev = document.createElement("span");
          cchev.className = "fs-cf-dialog__color-swatch-btn__chev";
          cchev.textContent = "\u25BE";
          cbtn.appendChild(cbar);
          cbtn.appendChild(cchev);
          const sync = (hex: string): void => {
            cbar.style.backgroundColor = hex;
          };
          return { btn: cbtn, sync };
        };

        const rowFillColors = document.createElement("div");
        rowFillColors.className = "fs-cf-dialog__row fs-cf-dialog__row--data-bar-fill-colors";
        const lPos = document.createElement("span");
        lPos.className = "fs-cf-dialog__label fs-cf-dialog__label--narrow";
        lPos.textContent = "正值:";
        const swPos = mkMiniSwatch();
        const lNeg = document.createElement("span");
        lNeg.className = "fs-cf-dialog__label fs-cf-dialog__label--narrow";
        lNeg.textContent = "负值:";
        const swNeg = mkMiniSwatch();
        rowFillColors.appendChild(lPos);
        rowFillColors.appendChild(swPos.btn);
        rowFillColors.appendChild(lNeg);
        rowFillColors.appendChild(swNeg.btn);
        secFill.appendChild(headFill);
        secFill.appendChild(rowFill);
        secFill.appendChild(rowFillColors);

        const sep2 = document.createElement("div");
        sep2.className = "fs-cf-dialog__scale-sep";

        const secBr = document.createElement("div");
        secBr.className = "fs-cf-dialog__cf-subsec";
        const headBr = document.createElement("div");
        headBr.className = "fs-cf-dialog__cf-subsec-title";
        headBr.textContent = "边框";
        const rowBr = document.createElement("div");
        rowBr.className = "fs-cf-dialog__row fs-cf-dialog__row--cf-radios";
        const borderSolid = document.createElement("input");
        borderSolid.type = "radio";
        borderSolid.name = "fsCfDbarBrKind";
        borderSolid.value = "solid";
        borderSolid.checked = true;
        const lblBs = document.createElement("label");
        lblBs.className = "fs-cf-dialog__inline-radio";
        lblBs.appendChild(borderSolid);
        lblBs.appendChild(document.createTextNode(" 实心边框 "));
        const borderNone = document.createElement("input");
        borderNone.type = "radio";
        borderNone.name = "fsCfDbarBrKind";
        borderNone.value = "none";
        const lblBn = document.createElement("label");
        lblBn.className = "fs-cf-dialog__inline-radio";
        lblBn.appendChild(borderNone);
        lblBn.appendChild(document.createTextNode(" 无边框 "));
        rowBr.appendChild(lblBs);
        rowBr.appendChild(lblBn);

        const rowBrColors = document.createElement("div");
        rowBrColors.className = "fs-cf-dialog__row fs-cf-dialog__row--data-bar-fill-colors";
        const lbPos = document.createElement("span");
        lbPos.className = "fs-cf-dialog__label fs-cf-dialog__label--narrow";
        lbPos.textContent = "正值:";
        const brPos = mkMiniSwatch();
        const lbNeg = document.createElement("span");
        lbNeg.className = "fs-cf-dialog__label fs-cf-dialog__label--narrow";
        lbNeg.textContent = "负值:";
        const brNeg = mkMiniSwatch();
        rowBrColors.appendChild(lbPos);
        rowBrColors.appendChild(brPos.btn);
        rowBrColors.appendChild(lbNeg);
        rowBrColors.appendChild(brNeg.btn);
        secBr.appendChild(headBr);
        secBr.appendChild(rowBr);
        secBr.appendChild(rowBrColors);

        const sep3 = document.createElement("div");
        sep3.className = "fs-cf-dialog__scale-sep";

        const rowAxis = document.createElement("div");
        rowAxis.className = "fs-cf-dialog__row fs-cf-dialog__row--data-bar-axis";
        const axLab = document.createElement("span");
        axLab.className = "fs-cf-dialog__label";
        axLab.textContent = "坐标轴位置:";
        const selAxis = document.createElement("select");
        selAxis.className = "fs-cf-dialog__select fs-cf-dialog__select--grow";
        for (const o of [
          { v: "automatic" as CfDataBarAxisPosition, t: "自动" },
          { v: "midpoint" as CfDataBarAxisPosition, t: "中间值" },
          { v: "none" as CfDataBarAxisPosition, t: "无" },
        ]) {
          const op = document.createElement("option");
          op.value = o.v;
          op.textContent = o.t;
          selAxis.appendChild(op);
        }
        rowAxis.appendChild(axLab);
        rowAxis.appendChild(selAxis);

        const rowAxisColor = document.createElement("div");
        rowAxisColor.className = "fs-cf-dialog__row fs-cf-dialog__row--data-bar-axis";
        const axColLab = document.createElement("span");
        axColLab.className = "fs-cf-dialog__label";
        axColLab.textContent = "坐标轴颜色:";
        const cAxis = mkMiniSwatch();
        rowAxisColor.appendChild(axColLab);
        rowAxisColor.appendChild(cAxis.btn);

        scaleCols.appendChild(sMin.col);
        scaleCols.appendChild(sMax.col);
        wrap.appendChild(scaleSep0);
        wrap.appendChild(scaleCols);
        wrap.appendChild(rowDir);
        wrap.appendChild(sep1);
        wrap.appendChild(secFill);
        wrap.appendChild(sep2);
        wrap.appendChild(secBr);
        wrap.appendChild(sep3);
        wrap.appendChild(rowAxis);
        wrap.appendChild(rowAxisColor);
        styleHost.appendChild(wrap);

        const dbLive: CfNewRuleLive = {
          k: "dataBar",
          sMin,
          sMax,
          selDir,
          chkShowBarOnly,
          fillSolid,
          fillGrad,
          borderSolid,
          borderNone,
          cFillPos: swPos.btn,
          syncFillPos: swPos.sync,
          cFillNeg: swNeg.btn,
          syncFillNeg: swNeg.sync,
          cBrPos: brPos.btn,
          syncBrPos: brPos.sync,
          cBrNeg: brNeg.btn,
          syncBrNeg: brNeg.sync,
          selAxis,
          cAxis: cAxis.btn,
          syncAxis: cAxis.sync,
          posFillHex: "#638ec6",
          negFillHex: "#ff0000",
          posBorderHex: "#2f5597",
          negBorderHex: "#000000",
          axisHex: "#000000",
          syncDataBarValueUi: (): void => {
            const tmin = sMin.st.value as CfDataBarMinEndpointType;
            const tmax = sMax.st.value as CfDataBarMaxEndpointType;
            if (!cfDataBarMinNeedsValue(tmin)) {
              sMin.inp.disabled = true;
              sMin.inp.value = "";
              sMin.inp.placeholder =
                tmin === "automatic" ? "（自动）" : tmin === "lowest" ? "（最低值）" : "";
            } else {
              sMin.inp.disabled = false;
              sMin.inp.placeholder = "";
            }
            if (!cfDataBarMaxNeedsValue(tmax)) {
              sMax.inp.disabled = true;
              sMax.inp.value = "";
              sMax.inp.placeholder =
                tmax === "automatic" ? "（自动）" : tmax === "highest" ? "（最高值）" : "";
            } else {
              sMax.inp.disabled = false;
              sMax.inp.placeholder = "";
            }
          },
        };

        dbLive.syncFillPos(dbLive.posFillHex);
        dbLive.syncFillNeg(dbLive.negFillHex);
        dbLive.syncBrPos(dbLive.posBorderHex);
        dbLive.syncBrNeg(dbLive.negBorderHex);
        dbLive.syncAxis(dbLive.axisHex);

        const refreshBrUi = (): void => {
          const en = borderSolid.checked;
          brPos.btn.disabled = !en;
          brNeg.btn.disabled = !en;
          brPos.btn.style.opacity = en ? "1" : "0.45";
          brNeg.btn.style.opacity = en ? "1" : "0.45";
        };

        const wirePick = (
          btn: HTMLButtonElement,
          get: () => string,
          set: (h: string) => void,
          sync: (h: string) => void,
        ): void => {
          btn.addEventListener("click", () => {
            scaleCloseColorPop?.();
            scaleCloseColorPop = mountCfColorPopover(btn, get(), (h) => {
              set(h);
              sync(h);
            });
          });
        };

        wirePick(
          dbLive.cFillPos,
          () => dbLive.posFillHex,
          (h) => {
            dbLive.posFillHex = h;
          },
          dbLive.syncFillPos,
        );
        wirePick(
          dbLive.cFillNeg,
          () => dbLive.negFillHex,
          (h) => {
            dbLive.negFillHex = h;
          },
          dbLive.syncFillNeg,
        );
        wirePick(
          dbLive.cBrPos,
          () => dbLive.posBorderHex,
          (h) => {
            dbLive.posBorderHex = h;
          },
          dbLive.syncBrPos,
        );
        wirePick(
          dbLive.cBrNeg,
          () => dbLive.negBorderHex,
          (h) => {
            dbLive.negBorderHex = h;
          },
          dbLive.syncBrNeg,
        );
        wirePick(
          dbLive.cAxis,
          () => dbLive.axisHex,
          (h) => {
            dbLive.axisHex = h;
          },
          dbLive.syncAxis,
        );

        sMin.st.addEventListener("change", () => {
          dbLive.syncDataBarValueUi();
          refreshOk();
        });
        sMax.st.addEventListener("change", () => {
          dbLive.syncDataBarValueUi();
          refreshOk();
        });
        sMin.inp.addEventListener("input", refreshOk);
        sMax.inp.addEventListener("input", refreshOk);
        selDir.addEventListener("change", refreshOk);
        chkShowBarOnly.addEventListener("change", refreshOk);
        fillSolid.addEventListener("change", refreshOk);
        fillGrad.addEventListener("change", refreshOk);
        borderSolid.addEventListener("change", () => {
          refreshBrUi();
          refreshOk();
        });
        borderNone.addEventListener("change", () => {
          refreshBrUi();
          refreshOk();
        });
        selAxis.addEventListener("change", refreshOk);

        live = dbLive;
        dbLive.syncDataBarValueUi();
        refreshBrUi();
        return dbLive;
      }
      return null;
    };

    selUiFamily.addEventListener("change", () => {
      const ui = selUiFamily.value as ConditionalFormatUiFamily;
      const unsupported = ui === "iconSet";
      tipNonClassic.hidden = !unsupported;
      paintStyleForm();
      refreshOk();
    });

    ok.addEventListener("click", () => {
      const rule = readRuleFromForm();
      if (rule === null) {
        window.alert(
          "无法应用当前规则：请选择「经典」「双色刻度」「三色刻度」或「数据条」，并填写完整条件。",
        );
        return;
      }
      finish(rule);
    });

    const initialForm = paintStyleForm();
    if (seedPartial !== null && initialForm?.k === "classic") {
      initialForm.selClassicType.value = seedPartial.classicType;
      initialForm.rebuildDynamic();
      initialForm.syncPreview();
    }
    refreshOk();

    const onKey = (ev: KeyboardEvent): void => {
      if (ev.key === "Escape") {
        finish(null);
      }
    };
    backdrop.tabIndex = -1;
    backdrop.addEventListener("keydown", onKey);
    backdrop.addEventListener("click", (ev) => {
      if (ev.target === backdrop) {
        finish(null);
      }
    });
    let teardown: (() => void) | undefined = () => {
      disposeStyleForm();
      backdrop.removeEventListener("keydown", onKey);
      backdrop.remove();
    };

    requestAnimationFrame(() => {
      backdrop.focus();
      ok.focus();
    });
  });
}

/** 简单「管理规则」：列出规则、删除单项、确定写回。 */
export function showConditionalFormatManageRulesDialog(
  initialRules: readonly ConditionalFormatRule[],
  sheetLabel: string,
): Promise<ConditionalFormatRule[] | null> {
  return new Promise((resolve) => {
    let settled = false;
    const finish = (v: ConditionalFormatRule[] | null): void => {
      if (settled) {
        return;
      }
      settled = true;
      teardown?.();
      resolve(v);
    };

    let rules = [...initialRules];

    const backdrop = document.createElement("div");
    backdrop.className = "fs-cf-dialog-backdrop";
    if (ribbonIsDarkTheme()) {
      backdrop.classList.add("fs-cf-dialog-backdrop--dark");
    }
    const dlg = document.createElement("div");
    dlg.className = "fs-cf-dialog fs-cf-dialog--manage";
    dlg.setAttribute("role", "dialog");
    dlg.setAttribute("aria-modal", "true");

    const header = document.createElement("div");
    header.className = "fs-cf-dialog__header";
    const title = document.createElement("div");
    title.className = "fs-cf-dialog__title";
    title.textContent = `条件格式规则 — ${sheetLabel}`;
    const closeBtn = document.createElement("button");
    closeBtn.type = "button";
    closeBtn.className = "fs-cf-dialog__close";
    closeBtn.setAttribute("aria-label", "关闭");
    closeBtn.textContent = "\u00D7";
    closeBtn.addEventListener("click", () => finish(null));
    header.appendChild(title);
    header.appendChild(closeBtn);

    const body = document.createElement("div");
    body.className = "fs-cf-dialog__body";
    const list = document.createElement("ul");
    list.className = "fs-cf-manage-list";

    const renderList = (): void => {
      list.textContent = "";
      if (rules.length === 0) {
        const li = document.createElement("li");
        li.className = "fs-cf-manage-empty";
        li.textContent = "当前工作表没有条件格式规则。";
        list.appendChild(li);
        return;
      }
      for (const r of rules) {
        const li = document.createElement("li");
        li.className = "fs-cf-manage-item";
        const lab = document.createElement("span");
        lab.className = "fs-cf-manage-item__label";
        lab.textContent = summarizeRule(r);
        const del = document.createElement("button");
        del.type = "button";
        del.className = "fs-cf-manage-item__del";
        del.textContent = "删除";
        del.addEventListener("click", () => {
          rules = rules.filter((x) => x.id !== r.id);
          renderList();
        });
        li.appendChild(lab);
        li.appendChild(del);
        list.appendChild(li);
      }
    };

    renderList();
    body.appendChild(list);

    const footer = document.createElement("div");
    footer.className = "fs-cf-dialog__footer";
    const cancel = document.createElement("button");
    cancel.type = "button";
    cancel.className = "fs-cf-dialog__btn fs-cf-dialog__btn--secondary";
    cancel.textContent = "取消";
    cancel.addEventListener("click", () => finish(null));
    const ok = document.createElement("button");
    ok.type = "button";
    ok.className = "fs-cf-dialog__btn fs-cf-dialog__btn--primary";
    ok.textContent = "确定";
    ok.addEventListener("click", () => finish(rules));
    footer.appendChild(cancel);
    footer.appendChild(ok);

    dlg.appendChild(header);
    dlg.appendChild(body);
    dlg.appendChild(footer);
    backdrop.appendChild(dlg);
    document.body.appendChild(backdrop);

    let teardown: (() => void) | undefined = () => {
      backdrop.remove();
    };
    const onKey = (ev: KeyboardEvent): void => {
      if (ev.key === "Escape") {
        finish(null);
      }
    };
    backdrop.tabIndex = -1;
    backdrop.addEventListener("keydown", onKey);
    requestAnimationFrame(() => {
      backdrop.focus();
    });
  });
}

function summarizeRule(r: ConditionalFormatRule): string {
  const rg = `R${r.range.startRow + 1}:C${r.range.startCol + 1}`;
  if (r.uiFamily === "twoColorScale") {
    return `${rg} · 双色刻度`;
  }
  if (r.uiFamily === "threeColorScale") {
    return `${rg} · 三色刻度`;
  }
  if (r.uiFamily === "dataBar") {
    return `${rg} · 数据条`;
  }
  if (r.classicType === "cellsThatContain") {
    return `${rg} · ${r.cellsThatContainKind ?? "cellValue"}`;
  }
  if (r.classicType === "topBottomRanked") {
    return `${rg} · 排名 ${r.topBottomKind ?? "top"}`;
  }
  if (r.classicType === "aboveBelowAverage") {
    return `${rg} · 平均值`;
  }
  if (r.classicType === "uniqueOrDuplicate") {
    return `${rg} · ${r.uniqueKind === "unique" ? "唯一" : "重复"}`;
  }
  return `${rg} · 规则`;
}
