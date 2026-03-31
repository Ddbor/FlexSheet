---
name: flexsheet-formula-test
description: Defines Vitest strategies for FlexSheet formula tests—parse correctness, SUM/AVERAGE/IF/AND/OR/VLOOKUP, A1 and cross-sheet refs, auto-recalc and dirty propagation, circular reference detection, #DIV/0! and invalid formulas. Use when writing or reviewing formula unit tests, 公式测试、单元格引用、循环引用、脏值、自动计算、非法公式.
---

# FlexSheet 公式测试（formula-test）

**版本**: 1.0

与 [flexsheet-test](../flexsheet-test/SKILL.md)（Vitest、目录、`*.test.ts`、pnpm 等）及 [flexsheet-formula-engine](../flexsheet-formula-engine/SKILL.md)（解析器/计算器契约）对齐：本技能只约束**公式相关**单测的范围与断言要点。

## 框架与放置

- 测试框架、目录与文件命名**完全遵循** [flexsheet-test](../flexsheet-test/SKILL.md)；公式用例可放在 `test/formula/` 等子目录。
- 优先通过**公开 API**（如 `Workbook` / 工作表设公式、读计算结果，或导出的 `parseFormula` / `evaluate`）断言；避免耦合未导出的内部符号，除非模块明确导出供稳定契约测试。

## 规则 1：公式解析正确性

- **正向**：合法公式（含嵌套括号、运算符优先级、字符串与数字字面量）应解析为可遍历的 AST（或等价结构），无静默失败。
- **负向**：明显非法语法（未闭合括号、非法 token）应得到**可定位**的错误或约定错误类型，而非随机数值。
- **断言**：至少覆盖「简单表达式」与「一层嵌套函数调用」两类；若项目支持，可加 R1C1 与 A1 对照用例。

## 规则 2：基础内置函数

对 **SUM、AVERAGE（或 AVG）、IF、AND、OR、VLOOKUP** 各至少：

| 函数 | 建议断言点 |
|------|------------|
| SUM | 多格区域、空单元格是否按产品约定参与（通常空作 0）、单参与多参 |
| AVERAGE | 分母为参与计数的数值个数；全空或全文本时的错误/忽略行为与引擎一致 |
| IF | 真/假分支、短路语义（与 [flexsheet-formula-engine](../flexsheet-formula-engine/SKILL.md) 中 IF 约定一致） |
| AND / OR | 全真、全假、混合、与空/错误混用时的约定 |
| VLOOKUP | 精确匹配、未找到 `#N/A`、列索引越界、区域非矩形（若可构造） |

函数名**大小写不敏感**时，至少一条用例混用大小写验证。

## 规则 3：A1 引用与跨 sheet

- **A1**：单列、多列字母、绝对引用 `$A$1` / 混合 `$A1`，区域 `A1:B2`。
- **跨表**：`Sheet1!A1` 或项目约定的引号表名形式；切换活动表或按表索引取值后结果仍正确。
- **断言**：依赖变更后，被引用单元格更新则公式结果随之更新（与规则 4 可合并用例）。

## 规则 4：自动计算与脏值

- **自动计算**：修改某单元格输入后，依赖该格的公式单元格在**同一次或下一轮**重算后得到新结果（行为以产品 API 为准：同步或微任务）。
- **脏值 / 依赖**：若 API 暴露「脏集合」或「需重算」标记，断言从叶子到根的更新顺序或最终一致性；至少验证**链式依赖**（A→B→C）全链更新正确。

## 规则 5：错误与非法公式

- **除零**：产生引擎约定的 **`#DIV/0!`**（或项目等价枚举），不得抛未捕获异常或返回普通数字。
- **非法公式**：无法解析、引用不存在的表/区域（若可模拟）时返回 **`#REF!`** 或约定错误，与 [flexsheet-formula-engine](../flexsheet-formula-engine/SKILL.md) 错误模型一致。
- **循环引用**：构造 A1 依赖 B1、B1 依赖 A1（或更长环），应被**检测**并返回约定错误/迭代策略结果，而非死循环或静默错误值。

## 新增公式测试自检

- [ ] 解析正例 + 至少一类解析失败路径
- [ ] SUM / AVERAGE / IF / AND / OR / VLOOKUP 均有代表用例
- [ ] A1 + 至少一种跨 sheet 形式
- [ ] 依赖链或脏值路径有一条可观测断言
- [ ] 除零、非法公式、循环引用各至少一条

## 延伸阅读

- 通用测试规范：[flexsheet-test](../flexsheet-test/SKILL.md)
- 解析与求值契约：[flexsheet-formula-engine](../flexsheet-formula-engine/SKILL.md)
- 解析细节（若需）：[flexsheet-formula-parser](../flexsheet-formula-parser/SKILL.md)
