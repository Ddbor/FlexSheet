import {
  normalizeSelectionRange,
  type ICommand,
  type SelectionRange,
  type Worksheet,
} from "@flexsheet/core";
import { FillSeriesCommand, type FillSeriesOptions } from "../commands/fill-series-command.js";
import { ensureFsSheetPromptStyles } from "./fs-dialog-styles.js";

let fillSeriesStylesInjected = false;

function ensureFillSeriesDialogStyles(): void {
  ensureFsSheetPromptStyles();
  if (fillSeriesStylesInjected) {
    return;
  }
  fillSeriesStylesInjected = true;
  const style = document.createElement("style");
  style.setAttribute("data-fs-fill-series-dialog", "1");
  style.textContent = `
.fs-fill-series.fs-sheet-prompt {
  width: min(520px, calc(100vw - 32px));
}
.fs-fill-series__body.fs-sheet-prompt__body {
  padding-top: 10px;
}
.fs-fill-series__top {
  display: flex;
  gap: 14px;
  align-items: stretch;
}
.fs-fill-series__col {
  flex: 1 1 0;
  min-width: 0;
}
.fs-fill-series__col + .fs-fill-series__col {
  border-left: 1px solid #e1dfdd;
  padding-left: 14px;
}
.fs-fill-series__col-title {
  margin: 0 0 6px 0;
  font-size: 13px;
  font-weight: 600;
  color: #323130;
}
.fs-fill-series__opt {
  display: flex;
  align-items: center;
  gap: 7px;
  margin: 5px 0;
  color: #323130;
  font-size: 13px;
}
.fs-fill-series__opt input {
  margin: 0;
}
.fs-fill-series__opt--disabled {
  color: #a19f9d;
}
.fs-fill-series__fields {
  margin-top: 12px;
  padding-top: 10px;
  border-top: 1px solid #e1dfdd;
}
.fs-fill-series__field {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-top: 8px;
}
.fs-fill-series__field:first-child {
  margin-top: 0;
}
.fs-fill-series__field-label {
  width: 52px;
  flex-shrink: 0;
  color: #323130;
  font-size: 13px;
}
.fs-fill-series__field .fs-sheet-prompt__input {
  width: 160px;
  padding-top: 3px;
  padding-bottom: 3px;
}
`;
  document.head.appendChild(style);
}

export interface FillSeriesDialogHost {
  readonly workbook: { getActiveSheet(): Worksheet | undefined } | undefined;
  readonly selection: { getNormalizedRange(): SelectionRange };
  readonly workspace: { readonly commands: { execute(cmd: ICommand): void } };
  refresh(): void;
}

function parseFiniteInput(input: HTMLInputElement): number | null {
  const t = input.value.trim();
  if (t === "") {
    return null;
  }
  const n = Number(t);
  return Number.isFinite(n) ? n : null;
}

/** Ribbon「填充 -> 系列」对话框。 */
export function showFillSeriesDialog(host: FillSeriesDialogHost): void {
  const sheet = host.workbook?.getActiveSheet();
  if (sheet === undefined) {
    return;
  }
  const range = normalizeSelectionRange(host.selection.getNormalizedRange());

  ensureFillSeriesDialogStyles();
  const overlay = document.createElement("div");
  overlay.className = "fs-sheet-prompt-overlay";
  overlay.setAttribute("role", "presentation");

  const panel = document.createElement("div");
  panel.className = "fs-sheet-prompt fs-fill-series";
  panel.setAttribute("role", "dialog");
  panel.setAttribute("aria-modal", "true");
  panel.setAttribute("aria-labelledby", "fs-fill-series-title");

  const header = document.createElement("div");
  header.className = "fs-sheet-prompt__header";
  const title = document.createElement("div");
  title.id = "fs-fill-series-title";
  title.className = "fs-sheet-prompt__title";
  title.textContent = "系列";
  const closeBtn = document.createElement("button");
  closeBtn.type = "button";
  closeBtn.className = "fs-sheet-prompt__close";
  closeBtn.setAttribute("aria-label", "关闭");
  closeBtn.textContent = "×";
  header.appendChild(title);
  header.appendChild(closeBtn);

  const body = document.createElement("div");
  body.className = "fs-sheet-prompt__body fs-fill-series__body";

  const top = document.createElement("div");
  top.className = "fs-fill-series__top";
  body.appendChild(top);

  const makeCol = (titleText: string): HTMLDivElement => {
    const col = document.createElement("div");
    col.className = "fs-fill-series__col";
    const h = document.createElement("h4");
    h.className = "fs-fill-series__col-title";
    h.textContent = titleText;
    col.appendChild(h);
    return col;
  };
  const seriesCol = makeCol("系列产生在:");
  const typeCol = makeCol("类型:");
  const dateCol = makeCol("日期单位:");
  top.appendChild(seriesCol);
  top.appendChild(typeCol);
  top.appendChild(dateCol);

  const makeRadio = (
    parent: HTMLElement,
    groupName: string,
    value: string,
    labelText: string,
    checked: boolean,
  ): HTMLInputElement => {
    const row = document.createElement("label");
    row.className = "fs-fill-series__opt";
    const ip = document.createElement("input");
    ip.type = "radio";
    ip.name = groupName;
    ip.value = value;
    ip.checked = checked;
    const text = document.createElement("span");
    text.textContent = labelText;
    row.appendChild(ip);
    row.appendChild(text);
    parent.appendChild(row);
    return ip;
  };
  const uniq = `${Date.now()}-${Math.random().toString(36).slice(2, 7)}`;
  const seriesGroup = `fs-fill-series-in-${uniq}`;
  const typeGroup = `fs-fill-series-type-${uniq}`;
  const unitGroup = `fs-fill-series-unit-${uniq}`;

  const seriesRow = makeRadio(seriesCol, seriesGroup, "row", "行", true);
  makeRadio(seriesCol, seriesGroup, "col", "列", false);

  const typeLinear = makeRadio(typeCol, typeGroup, "linear", "线性", true);
  const typeGrowth = makeRadio(typeCol, typeGroup, "growth", "等比序列", false);
  const typeDate = makeRadio(typeCol, typeGroup, "date", "日期", false);
  const typeAutofill = makeRadio(typeCol, typeGroup, "autofill", "自动填充", false);

  const makeUnit = (value: string, label: string, checked: boolean): HTMLInputElement => {
    return makeRadio(dateCol, unitGroup, value, label, checked);
  };
  const unitDay = makeUnit("day", "天", true);
  const unitWeekday = makeUnit("weekday", "工作日", false);
  const unitMonth = makeUnit("month", "月", false);
  const unitYear = makeUnit("year", "年", false);

  const trendRow = document.createElement("label");
  trendRow.className = "fs-fill-series__opt";
  const trend = document.createElement("input");
  trend.type = "checkbox";
  const trendText = document.createElement("span");
  trendText.textContent = "预测趋势";
  trendRow.appendChild(trend);
  trendRow.appendChild(trendText);
  typeCol.appendChild(trendRow);

  const fields = document.createElement("div");
  fields.className = "fs-fill-series__fields";
  const stepField = document.createElement("label");
  stepField.className = "fs-fill-series__field";
  const stepLabel = document.createElement("span");
  stepLabel.className = "fs-fill-series__field-label";
  stepLabel.textContent = "步长值:";
  const stepInput = document.createElement("input");
  stepInput.className = "fs-sheet-prompt__input";
  stepInput.type = "text";
  stepInput.value = "1";
  stepInput.setAttribute("autocomplete", "off");
  stepField.appendChild(stepLabel);
  stepField.appendChild(stepInput);

  const stopField = document.createElement("label");
  stopField.className = "fs-fill-series__field";
  const stopLabel = document.createElement("span");
  stopLabel.className = "fs-fill-series__field-label";
  stopLabel.textContent = "终止值:";
  const stopInput = document.createElement("input");
  stopInput.className = "fs-sheet-prompt__input";
  stopInput.type = "text";
  stopInput.value = "";
  stopInput.setAttribute("autocomplete", "off");
  stopField.appendChild(stopLabel);
  stopField.appendChild(stopInput);

  fields.appendChild(stepField);
  fields.appendChild(stopField);
  body.appendChild(fields);

  const footer = document.createElement("div");
  footer.className = "fs-sheet-prompt__footer";
  const okBtn = document.createElement("button");
  okBtn.type = "button";
  okBtn.className = "fs-sheet-prompt__btn fs-sheet-prompt__btn--primary";
  okBtn.textContent = "确定";
  const cancelBtn = document.createElement("button");
  cancelBtn.type = "button";
  cancelBtn.className = "fs-sheet-prompt__btn fs-sheet-prompt__btn--secondary";
  cancelBtn.textContent = "取消";
  footer.appendChild(cancelBtn);
  footer.appendChild(okBtn);

  panel.appendChild(header);
  panel.appendChild(body);
  panel.appendChild(footer);
  overlay.appendChild(panel);
  document.body.appendChild(overlay);

  const readType = (): FillSeriesOptions["type"] => {
    if (typeGrowth.checked) {
      return "growth";
    }
    if (typeDate.checked) {
      return "date";
    }
    if (typeAutofill.checked) {
      return "autofill";
    }
    return "linear";
  };

  const syncDateUnitsEnabled = (): void => {
    const enabled = typeDate.checked;
    for (const ip of [unitDay, unitWeekday, unitMonth, unitYear]) {
      ip.disabled = !enabled;
      const row = ip.closest(".fs-fill-series__opt");
      if (row !== null) {
        row.classList.toggle("fs-fill-series__opt--disabled", !enabled);
      }
    }
  };

  const remove = (): void => {
    document.removeEventListener("keydown", onKey, true);
    overlay.remove();
  };

  const onKey = (e: KeyboardEvent): void => {
    if (e.key === "Escape") {
      e.preventDefault();
      remove();
    }
  };

  const confirm = (): void => {
    const type = readType();
    const parsedStep = parseFiniteInput(stepInput);
    if (type !== "autofill" && parsedStep === null) {
      stepInput.focus();
      stepInput.select();
      return;
    }
    const stop = parseFiniteInput(stopInput);
    const dateUnit: FillSeriesOptions["dateUnit"] = unitWeekday.checked
      ? "weekday"
      : unitMonth.checked
        ? "month"
        : unitYear.checked
          ? "year"
          : "day";
    const options: FillSeriesOptions = {
      seriesIn: seriesRow.checked ? "row" : "col",
      type,
      dateUnit,
      step: parsedStep ?? 1,
      stop,
      trend: trend.checked,
    };
    host.workspace.commands.execute(new FillSeriesCommand(sheet, range, options));
    host.refresh();
    remove();
  };

  closeBtn.addEventListener("click", remove);
  cancelBtn.addEventListener("click", remove);
  okBtn.addEventListener("click", confirm);
  overlay.addEventListener("click", (e) => {
    if (e.target === overlay) {
      remove();
    }
  });
  for (const ip of [typeLinear, typeGrowth, typeDate, typeAutofill]) {
    ip.addEventListener("change", syncDateUnitsEnabled);
  }
  document.addEventListener("keydown", onKey, true);
  syncDateUnitsEnabled();
  queueMicrotask(() => {
    stepInput.focus();
    stepInput.select();
  });
}
