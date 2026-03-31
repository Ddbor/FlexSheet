---
name: flexsheet-formula-engine
description: Specifies FlexSheet formula engine responsibilities—FormulaParser, FormulaCalculator, built-ins (SUM, AVG, IF, VLOOKUP), and A1/R1C1 cell references. Use when implementing or reviewing formula parsing, evaluation, dependency tracking, or when the user mentions 公式、公式引擎、单元格引用、A1、R1C1、VLOOKUP, or src/formula.
---

# FlexSheet 公式引擎规范

与 [flexsheet-architecture](../flexsheet-architecture/SKILL.md) 分工：架构技能定义 **Formula 层** 与 `src/formula/` 目录归属；本技能定义**解析器、计算器、引用与常用函数**的生成与实现约束。实现细节须同时遵守 [flexsheet-ts-style](../flexsheet-ts-style/SKILL.md)。**词法/语法与 AST 形态**的专精约定见 [flexsheet-formula-parser](../flexsheet-formula-parser/SKILL.md)（本技能为求值与语义主文档）。

## 放置与边界

- 公式 AST、解析、求值、依赖收集放在 **`src/formula/`**（或等价物理目录），**不**混入 `render/`。
- **FormulaParser** 只负责语法 → AST；**FormulaCalculator** 只负责在「工作簿/工作表上下文」下对 AST 求值；二者通过清晰类型边界交互，避免在 Parser 里做数值计算。

## FormulaParser（解析器）

**职责**

- 输入：以 `=` 开头的公式字符串（或项目约定的公式单元格原始串）。
- 输出：**抽象语法树（AST）**，节点类型覆盖：字面量（数字、字符串、布尔）、单元格引用、区域引用、函数调用、运算符表达式、括号。
- 错误：词法/语法错误时返回**可定位**的诊断（偏移或行列），供 UI 标红；不静默吞错。

**须支持**

- 常见运算符优先级与结合性（与 Excel 类表格一致：算术、比较、字符串连接 `&` 等按产品定义实现）。
- **单元格与区域引用**（见下文 A1 / R1C1）。
- **函数调用**：名称大小写规则（通常不敏感）、参数列表、嵌套调用。

**不宜**

- 在 Parser 内查询单元格值或执行业务函数体（交给 Calculator + 函数注册表）。

## FormulaCalculator（计算器）

**职责**

- 输入：已解析的 AST + **求值上下文**（当前工作表、坐标、可选循环检测栈等）。
- 输出：单元格标量结果（数字、字符串、布尔、错误类型如 `#DIV/0!`、`#REF!` 等，与产品错误模型一致）或数组/区域结果（若产品支持动态数组，再扩展类型）。
- **依赖收集**：求值同一公式时，应能列出其静态或动态依赖的单元格/区域，供依赖图与重算调度使用（与架构中 Formula 层「依赖图」一致）。

**不宜**

- 重复实现词法/语法分析（应调用 Parser 或消费其 AST）。

## 内置函数（约定语义，实现须一致）

下列为**必须覆盖语义**的代表函数；名称对大小写不敏感，参数个数与类型错误应返回对应错误值。

| 函数 | 要点 |
|------|------|
| **SUM** | 对若干数字、单元格、区域求和；忽略空单元格逻辑按 Excel 惯例（空作 0 参与求和等，与 Data 层空值模型对齐）。 |
| **AVG** / **AVERAGE** | 区域数值平均；分母为**参与计数的数值个数**（非区域面积，除非与产品约定为区域面积）。 |
| **IF** | 三参数：条件、真分支、假分支；短路求值（假分支在条件为假时可不算，除非语言定义为全算）。 |
| **VLOOKUP** | 四参数：查找值、查找列区域、列索引、近似/精确匹配；列索引越界、`#N/A` 未找到等须返回规范错误类型；区域须为矩形。 |

其他函数通过**同一套注册机制**扩展：名称 → 参数校验 → 依赖传播规则。

## 单元格引用：A1 与 R1C1

**A1**

- 列字母（或多字母）+ 行号，如 `A1`、`$B$2`（绝对 `$` 规则与 Excel 一致）。
- 区域：`A1:B2`、跨表若支持则 `Sheet1!A1`（具体分隔符与引号规则按产品定义）。

**R1C1**

- 形式 `R[row]C[col]`，相对/绝对用方括号，如 `R[1]C[1]`、`R1C1`（行为与 Excel R1C1 一致）。
- Parser 应能识别当前工作簿设置下的引用风格，或同时支持两种并在 AST 中标注引用类型，便于展示与序列化互转。

**统一内部表示**

- AST 中引用建议归一为 **结构化坐标**（表索引、行、列、绝对性标志），避免在核心逻辑里长期保留「原始字符串」。
- 提供 **A1 ↔ R1C1 ↔ 内部坐标** 的转换工具函数，供编辑栏与存储格式共用。

## 与数据层协作

- 取值、设值、空单元格含义由 **Data** 层定义；Calculator 通过窄接口（如 `getCellValue(address)`、`getRangeValues(range)`）访问，不直接依赖 Canvas 或 DOM。
- 重算顺序由依赖图决定；循环引用检测与迭代策略在 Formula 层或调度模块中明确，不在本技能展开实现细节。

## 自检清单

- [ ] Parser / Calculator 类名与职责是否与上表一致、无职责泄漏？
- [ ] AST 是否覆盖引用、区域、函数与运算符，且可遍历做依赖收集？
- [ ] SUM、AVG、IF、VLOOKUP 语义与错误是否与上表及空值模型一致？
- [ ] A1 与 R1C1 是否解析为统一内部表示并有互转路径？
