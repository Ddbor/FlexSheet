---
name: flexsheet-spreadjs-like
description: Positions FlexSheet as SpreadJS-aligned in features, public API shape, and interaction UX while requiring 100% independent TypeScript + Canvas 2D implementation—no copied code, no reverse engineering, ECMA-376 XLSX, layered core/render/interaction/formula/feature/adapter. Use when designing APIs or UX like SpreadJS, comparing behavior to spreadsheet hosts, scoping Excel-like features, or when the user mentions SpreadJS、类 SpreadJS、对标 SpreadJS、独立自研、纯前端 Canvas Excel.
---

# FlexSheet：SpreadJS 级体验，自研实现

**版本：1.0**

## 产品定位

- **对齐**：功能范围、交互逻辑、用户体验、对外 API 风格参考 **SpreadJS**（命名习惯、配置结构、事件模型等「产品面」一致）。
- **实现**：代码 **100% 自研**，不依赖 SpreadJS 源码、不复制其代码、不反编译、不照搬内部类/私有方法命名或受版权保护的具体实现细节。

底层为 **自研 Canvas 渲染、自研数据模型、自研公式引擎**；存储与交换遵循 **Excel 公开标准（含 ECMA-376 XLSX）**，不使用私有协议或专有封闭格式。

## 核心规则

1. **体验与能力**：与 SpreadJS 同类产品在常用表格能力上对齐（见下方参考功能点）。
2. **API 与配置**：对外 API、选项、事件命名与组织方式 **参考 SpreadJS 风格**，便于从同类宿主迁移；实现为全新设计与自有类型，不映射其内部类型体系。
3. **原创性**：禁止直接复制其代码；禁止机械照搬内部类层次与私有 API；禁止以逆向或「洗稿」方式复现非公开实现。
4. **引擎**：渲染、数据、公式均为 **项目内自研模块**（分层见 [flexsheet-architecture](../flexsheet-architecture/SKILL.md)）。
5. **标准**：行为与文件格式以 **Excel 兼容与 OOXML 公开规范** 为准，不依赖厂商私有扩展作为主路径。
6. **操作习惯**：选区、编辑、拖拽填充、复制粘贴、冻结窗格、合并、样式、公式、**标准 XLSX** 导入导出等，与常见电子表格习惯一致。
7. **架构**：保持 **core / render / interaction / formula / feature / adapter**（及 data/api 等）边界；细节以架构技能为准。
8. **运行环境**：**纯前端、无后端**；核心框架无关，跨框架通过 **Adapter** 接入（见 [flexsheet-cross-framework](../flexsheet-cross-framework/SKILL.md)）。

## 参考功能点（实现清单导向）

实现与文档时可对照，**不要求 API 名称与 SpreadJS 逐字相同**，但能力与交互应可对标：

| 领域 | 要点 |
|------|------|
| 结构 | Workbook / Worksheet（或多表模型） |
| 样式 | 单元格样式系统（字体、对齐、边框、填充等） |
| 选区 | 选区模型、活动单元格、扩展选区 |
| 编辑 | 单元格编辑、公式栏联动、编辑态与提交规则 |
| 交互 | 拖拽、填充柄、行列宽高调校 |
| 历史 | 撤销 / 重做 |
| 数据 | 数据验证 |
| 可视化 | 条件格式 |
| 图表 | 图表（Canvas 绘制与数据绑定，见 [flexsheet-chart](../flexsheet-chart/SKILL.md)） |
| 文件 | XLSX 导入导出（**ECMA-376**，见 [flexsheet-import-export](../flexsheet-import-export/SKILL.md)） |

更细子域请按需引用：`flexsheet-interaction`、`flexsheet-formula-engine`、`flexsheet-conditional-format`、`flexsheet-theme` 等。

## 禁止行为（红线）

- 直接复制 SpreadJS 或第三方受版权保护的源码片段。
- 照搬其 **内部** 类结构、私有方法名、非公开 API 作为本库内部实现依据。
- 反编译、刻意模仿其 **内部** 优化技巧或闭源行为以「对齐性能」为名侵权。
- 将任何 **受版权保护** 的逻辑原样迁入本仓库。

**允许**：阅读公开文档、观察 **用户可见** 行为与 Excel 标准行为，用自研代码与自有数据结构实现等效能力。

## 与其他技能的关系

| 技能 | 用途 |
|------|------|
| [flexsheet-architecture](../flexsheet-architecture/SKILL.md) | 分层与目录 |
| [flexsheet-ts-style](../flexsheet-ts-style/SKILL.md) | TypeScript 约定 |
| 各子域技能 | 渲染、交互、公式、导入导出、主题等 |

## Agent 执行要点

- 用户要求「像 SpreadJS」时：优先满足 **能力与交互约定**，API 可 **风格相近** 但类型与实现保持 FlexSheet 自有命名与分层。
- 任何新功能先判断是否落在 Data / Render / Interaction / Formula / Feature / Adapter 之一，避免跨层混杂。
- 涉及表格竞品时：强调 **标准与自研**，不建议引用或复制闭源实现细节。
