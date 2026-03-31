---
name: flexsheet-conditional-format
description: Implements Excel-style conditional formatting—data bars, color scales, icon sets, and rule-based cell highlighting—evaluated in Data/Feature and painted on Canvas 2D; Adapter maps or exports compatible rules. Pure frontend, no formula evaluation inside Render. Use when implementing or reviewing 条件格式、数据条、色阶、图标集、单元格高亮、CF、Excel 导出与条件格式兼容。
---

# FlexSheet 条件格式（Conditional Format）

**Version:** 1.0

与 [flexsheet-architecture](../flexsheet-architecture/SKILL.md)、[flexsheet-cell-render](../flexsheet-cell-render/SKILL.md)、[flexsheet-canvas-render](../flexsheet-canvas-render/SKILL.md) 一致；TypeScript 规范见 [flexsheet-ts-style](../flexsheet-ts-style/SKILL.md)。Adapter 与多格式构建见 [flexsheet-cross-framework](../flexsheet-cross-framework/SKILL.md)。

## 硬性规则

| 规则 | 说明 |
|------|------|
| **纯前端** | 无服务端；规则存储与求值均在浏览器内完成。 |
| **Canvas 渲染** | 数据条、色阶填充、图标、高亮叠加层均在 **Render** 用 `CanvasRenderingContext2D` 绘制；主表不用 DOM 矩阵。 |
| **不在 Render 算公式** | 条件是否成立、条/色阶的归一化数值、图标档位等，在 **Data / Feature** 求值后，以**只读结果对象**交给 `CellRender` / `StyleRender` 绘制（与 canvas-render 一致）。 |
| **导出兼容 Excel** | 通过 **Adapter** 输出 `.xlsx`（或约定格式）时，应映射 Open XML 条件格式语义，或明确文档化「降级策略」（例如仅导出最终填充色）。禁止静默丢失用户可见规则而不说明。 |

## 能力范围（四类）

| 类型 | 要点 |
|------|------|
| **数据条（Data bars）** | 在单元格内按比例绘制横向条（渐变或实色）；最小/最大值可来自区域、固定值或自动（区域 min/max）。条在**背景之上、文本之下或按产品约定分层**，避免挡字时调透明度或内边距。 |
| **色阶（Color scales）** | 二色或三色渐变；按单元格数值在 `[min, max]` 上线性映射到颜色。min/max 来源与数据条类似，需在模型层统一为一种「标度规范」便于测试与导出。 |
| **图标集（Icon sets）** | 按阈值或百分比分档显示图标；图标为**矢量路径或预打包位图**，DPR 下清晰；布局常见为条左/文右或图标列固定宽度，与对齐设置一致。 |
| **单元格高亮（规则）** | 基于比较运算符、文本包含、重复值、公式为真等规则，合并出**字体色 / 填充色 / 边框**等到 `CellStyle` 或扩展结构；多规则时定义**优先级与停止条件**（与 Excel 行为对齐或显式差异说明）。 |

## 分层职责

| 层 | 职责 |
|----|------|
| **Data** | 持久化规则列表（作用范围、类型、阈值、引用区域 ID）；单元格原始值供求值。 |
| **Feature** | 规则求值管线：依赖区域变更时增量或批量重算「有效显示样式 / CF 叠加层描述」。 |
| **Render** | 只读消费：先画单元格底色与数据条/色阶底纹，再画边框与文字，最后画图标；**不**改写单元格值。 |
| **Adapter** | 导入/导出时与 Excel 条件格式子集互转；不实现的特性在变更日志或类型上标为 `unsupported`。 |

## 绘制顺序（建议）

在单格内：`save` → 背景（含色阶）→ **数据条**（若与背景同层则先算 clip 区域）→ 文本 → **图标** → `restore`。与全表顺序仍遵循 cell-render：**清屏 → 网格 → 单元格 → 覆盖层**。

## Excel 导出（兼容）要点

- **优先**：导出规则定义（`cfRule` 等价信息）+ 区域，便于在 Excel 中继续编辑。
- **兜底**：无法映射的规则导出为**最终样式快照**（静态填充/字体），并在内部标记为已栅格化。
- **图标集**：需与 Excel 档位语义一致或限制为项目支持的子集；否则导出图标为嵌入图片或略过并文档说明。

## 自检清单

- [ ] 条件格式求值是否在 Data/Feature，Render 仅根据结果绘制？
- [ ] 数据条/色阶/图标是否与虚拟滚动、合并单元格主从格协调（主格绘制、从属跳过或继承）？
- [ ] 多规则优先级是否与产品说明一致？
- [ ] Adapter 是否覆盖「规则导出 / 样式快照」至少一种路径？

## 延伸阅读

- 样式与 `StyleRender`： [flexsheet-cell-render](../flexsheet-cell-render/SKILL.md)
- 视口与样式数据流： [flexsheet-canvas-render](../flexsheet-canvas-render/SKILL.md)
