---
name: flexsheet-formula-parser
description: Excel-style formula tokenization and parsing to AST—parseFormula entry, AST node kinds, A1/range in AST, no cell reads in Parser. Evaluation, Calculator, deps, cyclic refs, R1C1, and built-in semantics are in flexsheet-formula-engine. Use when implementing parseFormula, tokenizer, AST nodes, or when the user mentions 公式解析、AST、词法语法、src/formula Parser 侧.
---

# FlexSheet 公式解析（Parser 专精）

与 [flexsheet-formula-engine](../flexsheet-formula-engine/SKILL.md) 分工：引擎技能定义 **FormulaParser + FormulaCalculator、内置函数语义、A1/R1C1、依赖图与重算**；**本技能仅约束词法/语法管线与 AST 形态**。类型见 [flexsheet-ts-style](../flexsheet-ts-style/SKILL.md)。

## 边界

- **纯前端**：词法/语法在 TypeScript 中实现；**不**引入未许可的重型第三方解析库（可递归下降或 Pratt；轻量工具仅限团队明确允许）。
- **Parser 只做字符串 → AST**：**不在 Parser 内**读单元格值、不执行 `SUM`/`IF` 等函数体（交给 Calculator + 注册表）。求值、依赖收集、循环引用、自动重算均见**引擎技能**，此处不重复。

## 输入与输出

**输入**：公式字符串；通常以 `=` 开头（可在入口统一剥离）。

**输出**：`ASTNode` 为根的树，至少覆盖：

| 类别 | 说明 |
|------|------|
| 字面量 | 数字、字符串、布尔 |
| 引用 | 单格 `A1`、区域 `A1:B5`；`$` 按 Excel 规则进入 AST |
| 二元/一元运算 | 算术、比较、字符串连接 `&` 等，优先级与结合性与产品/Excel 对齐 |
| 函数调用 | 标识符 + 参数列表、嵌套，如 `SUM(A1:B5)`、`IF(...)`、`VLOOKUP(...)` |

**错误**：词法/语法失败时返回**可定位**诊断（偏移或行列），供编辑栏标红；不静默返回「伪 AST」。

**公开入口**：对外保持 `parseFormula`（或等价）**单一入口**；实现上可分 `strip '='` → `tokenize` → `parse`。

```ts
function parseFormula(formula: string): ASTNode {
  // 具体节点枚举与字段名以项目 types 为准
}
```

## A1 与跨表（解析侧）

- 列字母（含 `AA`…）+ 行号；区域 `起:止`。
- AST 内归一为**结构化地址**（工作表、行、列、绝对性标志等），避免长期只用原始字符串做依赖分析。
- 跨表 `Sheet!A1` 若产品支持，在 AST 中显式标注表名/索引。

**R1C1**、A1↔R1C1↔内部坐标互转、以及 **IF 短路**等**求值语义**以 [flexsheet-formula-engine](../flexsheet-formula-engine/SKILL.md) 为准；Parser 只保证 `IF` 在 AST 中为三参数函数调用节点。

## 自检清单

- [ ] `parseFormula` 是否只产出 AST，不做数值计算与单元格读取？
- [ ] AST 是否可遍历以供 Calculator 做依赖收集与求值（细节见引擎技能）？
- [ ] 是否无未许可的重型第三方解析库？
- [ ] A1/区域是否解析为统一内部坐标表示？

## 延伸阅读

- 求值、函数语义、循环引用、重算：[flexsheet-formula-engine](../flexsheet-formula-engine/SKILL.md)
