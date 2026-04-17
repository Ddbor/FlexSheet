import {
  normalizeSelectionRange,
  selectionRangesIntersect,
  type ConditionalFormatRule,
  type ICommand,
  type SelectionRange,
  type Worksheet,
} from "@flexsheet/core";

/** 追加条件格式规则（可撤销）。 */
export class AddConditionalFormatRuleCommand implements ICommand {
  readonly id = "sheet.addConditionalFormatRule";
  readonly label = "新建条件格式规则";
  private readonly before: readonly ConditionalFormatRule[];

  constructor(
    private readonly sheet: Worksheet,
    private readonly rule: ConditionalFormatRule,
  ) {
    this.before = [...sheet.getConditionalFormatRules()];
  }

  execute(): void {
    this.sheet.addConditionalFormatRule(this.rule);
  }

  undo(): void {
    this.sheet.setConditionalFormatRules(this.before);
  }
}

/** 用快照替换条件格式规则列表（可撤销）。 */
export class SetConditionalFormatRulesCommand implements ICommand {
  readonly id = "sheet.setConditionalFormatRules";
  readonly label = "设置条件格式规则";
  private readonly before: readonly ConditionalFormatRule[];

  constructor(
    private readonly sheet: Worksheet,
    private readonly next: readonly ConditionalFormatRule[],
  ) {
    this.before = [...sheet.getConditionalFormatRules()];
  }

  execute(): void {
    this.sheet.setConditionalFormatRules(this.next);
  }

  undo(): void {
    this.sheet.setConditionalFormatRules(this.before);
  }
}

/** 清除与选区相交的条件格式规则（可撤销）。 */
export class ClearConditionalFormatRulesIntersectingCommand implements ICommand {
  readonly id = "sheet.clearConditionalFormatRulesIntersecting";
  readonly label = "清除条件格式规则（所选单元格）";
  private readonly before: readonly ConditionalFormatRule[];

  constructor(
    private readonly sheet: Worksheet,
    private readonly range: SelectionRange,
  ) {
    this.before = [...sheet.getConditionalFormatRules()];
  }

  execute(): void {
    const n = normalizeSelectionRange(this.range);
    const after = this.before.filter((r) => !selectionRangesIntersect(r.range, n));
    this.sheet.setConditionalFormatRules(after);
  }

  undo(): void {
    this.sheet.setConditionalFormatRules(this.before);
  }
}

/** 清除整张工作表条件格式规则（可撤销）。 */
export class ClearAllConditionalFormatRulesCommand implements ICommand {
  readonly id = "sheet.clearAllConditionalFormatRules";
  readonly label = "清除条件格式规则（整个工作表）";
  private readonly before: readonly ConditionalFormatRule[];

  constructor(private readonly sheet: Worksheet) {
    this.before = [...sheet.getConditionalFormatRules()];
  }

  execute(): void {
    this.sheet.setConditionalFormatRules([]);
  }

  undo(): void {
    this.sheet.setConditionalFormatRules(this.before);
  }
}
