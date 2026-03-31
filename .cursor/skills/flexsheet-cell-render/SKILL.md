---
name: flexsheet-cell-render
description: Canvas 2D 单元格级绘制分工—CellRender / StyleRender、文档坐标与网格线、行列尺寸、renderCell 级 API；视口/虚拟滚动/冻结/合并的通用约定见 flexsheet-canvas-render。 Use when implementing cell-level paint, grid layout, row/column sizing, high-performance sheet rendering, or when the user mentions 单元格渲染、样式绘制、CellRender、StyleRender.
---

# FlexSheet 单元格渲染（Cell Render）

**Version:** 1.0

与 [flexsheet-canvas-render](../flexsheet-canvas-render/SKILL.md) 配合：**视口裁剪、虚拟滚动、冻结窗格、合并单元格、全表绘制顺序、禁止 DOM 主表格**等一律以该技能为准。本技能只规定**单元格粒度**的职责拆分与坐标/网格细节。架构见 [flexsheet-architecture](../flexsheet-architecture/SKILL.md)，TS 见 [flexsheet-ts-style](../flexsheet-ts-style/SKILL.md)。

## 本技能独有：渲染类分工

名称与代码检索保持一致：

| 类 | 职责 |
|----|------|
| **`CanvasRenderer`** | 门面—DPR、画布尺寸、可见区计算、调度重绘、协调各 painter（与 canvas-render 一致）。 |
| **`CellRender`** | 单格几何与内容—背景、边框、文本测量与绘制、`renderCell` 级 API。 |
| **`StyleRender`** | 将 `CellStyle`（及条件格式**已求值结果**）映射到 `ctx`（`font` / `fillStyle` / `strokeStyle` / `lineWidth` / `textAlign` / `textBaseline`），供 CellRender 调用。 |

Render 层**不**现场算公式；样式与条件格式结果对象来自 Data/Feature，见 canvas-render 与 [flexsheet-conditional-format](../flexsheet-conditional-format/SKILL.md)。

## 坐标、网格线与行列尺寸

- **文档坐标**：单元格 `(row, col)` → 左上角 `(x, y)` 由**累积行高、列宽**（含隐藏/筛选时跳过）算出；冻结偏移与 canvas-render、Interaction **同一套**反算。
- **网格线**：与单元格边框分层，避免与 `strokeRect` 双倍线宽；策略在仓库内二选一并文档化。
- **行列宽高**：单一数据源（如 `RowColSize` / `defaultRowHeight`）；变更后标记脏区或全表重绘。

## 绘制顺序（全表）

与 canvas-render 一致：**清屏/背景 → 网格线（若独立层）→ 单元格背景与边框 → 文本 → 覆盖层（选区、填充柄、冻结分隔线）**。

## 示例：`renderCell`

仅示意背景与边框；完整应对齐 StyleRender 的字体/对齐设置。

```typescript
function renderCell(
  ctx: CanvasRenderingContext2D,
  cell: Cell,
  x: number,
  y: number,
  width: number,
  height: number
): void {
  ctx.fillStyle = cell.style.bgColor;
  ctx.fillRect(x, y, width, height);
  ctx.strokeRect(x, y, width, height);
}
```

## 自检清单

- [ ] `CellRender` / `StyleRender` 与 `CanvasRenderer` 职责是否清晰、无在 Render 层写业务数据？
- [ ] 合并与冻结是否与滚动、命中测试共用 canvas-render 约定的坐标系？
- [ ] 详见： [flexsheet-canvas-render](../flexsheet-canvas-render/SKILL.md)

## 延伸阅读

- [flexsheet-canvas-render](../flexsheet-canvas-render/SKILL.md)
