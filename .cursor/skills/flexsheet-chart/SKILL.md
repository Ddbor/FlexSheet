---
name: flexsheet-chart
description: Canvas 2D 图表渲染（柱状图、折线图、饼图、面积图），绑定单元格区域为数据源，容器自适应与 DPR，经 Adapter 导出到 Excel。Use when implementing or reviewing chart widgets, Canvas charts, 图表、柱状图、折线图、饼图、面积图, cell-bound series, responsive chart layout, or Excel chart export in FlexSheet.
---

# FlexSheet 图表（Chart）

**Version:** 1.0

架构与分层见 [flexsheet-architecture](../flexsheet-architecture/SKILL.md)；Canvas 通用约定见 [flexsheet-canvas-render](../flexsheet-canvas-render/SKILL.md)；TS 规范见 [flexsheet-ts-style](../flexsheet-ts-style/SKILL.md)。

## 硬性规则

1. **纯 Canvas 2D**：图表本体只用 `CanvasRenderingContext2D` 绘制；禁止用 SVG/DOM 节点矩阵充当图表主渲染（图例/工具提示若用少量浮动层，须极薄且不占数据层职责）。
2. **绑定单元格数据**：系列名、分类轴标签、数值系列均从 **Data 层** 的单元格区域或结构化快照读取（如 `RangeRef` → 解析为 `number | string | null`）；Render 不内嵌业务表格逻辑，仅消费只读 `ChartDataSource`（或等价接口）。
3. **自适应**：监听容器尺寸（`ResizeObserver` 或宿主传入 `width`/`height`），按 CSS 像素设置 `canvas` 尺寸并乘以 **DPR** 修正 `canvas.width/height`；重绘与坐标系以 CSS 像素为逻辑单位，与现有 `CanvasRenderer` 一致。
4. **导出到 Excel**：经 **Adapter** 写出 `.xlsx`，须符合 [flexsheet-import-export](../flexsheet-import-export/SKILL.md)（合法 OPC/ZIP、ECMA-376）。至少可将图表依赖的单元格区域导出为**表格数据**；若需内嵌 Excel 图表对象，仅在 Adapter 内对接**项目选定**的 OOXML/工具，核心不 `import` 具体 xlsx 实现细节。

## 分层放置

| 关注点 | 层 |
|--------|-----|
| 单元格区域 → 二维/序列数据 | Data + 薄 API |
| 图表几何、绘制、动画帧 | Render（可命名 `ChartRenderer` / `*ChartPainter`） |
| 插入图表、选区更新数据源 | Feature / Interaction 命令 |
| `.xlsx` 写出与格式 | `adapters/` |

## 支持的图表类型（约定）

| 类型 | 数据形态要点 |
|------|----------------|
| **柱状图** | 分类 × 一个或多个系列；柱宽与组距随绘图区宽度缩放 |
| **折线图** | 有序 X（分类或数值轴）；点与线抗锯齿、线宽一致 |
| **饼图** | 单系列占比；扇区从统一起点顺时针；和为 0 或非法时降级为空态提示 |
| **面积图** | 与折线类似，路径闭合到基线后 `fill` + `stroke`，注意透明度叠层 |

## 数据与坐标

- **输入**：至少包含 `categories: string[]`、`series: { name: string; values: number[] }[]`，长度一致；缺失单元格映射为 `null` 并在绘制前过滤或按策略视为 0（策略在 Data/API 层定死一处）。
- **输出几何**：统一 `padding`、`plotWidth/plotHeight`，刻度与网格线在 plot 区内；图例占位可预留固定条带或重叠策略（文档化一种默认）。

## 性能与重绘

- 数据或容器变化 → 标记脏 → `requestAnimationFrame` 单次绘制；大数据系列时避免每帧分配大数组。
- 不与主表网格同画布时，独立 `OffscreenCanvas` 或独立 `<canvas>`，避免与 `CanvasRenderer` 争抢同一上下文状态。

## 自检清单

- [ ] 图表主绘制路径是否无 SVG/DOM 矩阵？
- [ ] 数值与标签是否均来自 Data/API 抽象，而非 Render 内硬编码？
- [ ] DPR 与容器缩放是否正确？
- [ ] Excel 导出是否仅在 Adapter，核心无具体 xlsx 库 import？

## 延伸阅读

- 单元格与样式只读模型： [flexsheet-cell-render](../flexsheet-cell-render/SKILL.md)
- XLSX 约束： [flexsheet-import-export](../flexsheet-import-export/SKILL.md)
- 跨框架与打包： [flexsheet-cross-framework](../flexsheet-cross-framework/SKILL.md)
