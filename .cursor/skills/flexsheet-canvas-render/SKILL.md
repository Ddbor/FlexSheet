---
name: flexsheet-canvas-render
description: Specifies Canvas 2D rendering for FlexSheet grid—CanvasRenderer, viewport/virtual scroll, cell painting, styles, frozen panes. Use when implementing or reviewing render code, drawing cells, scroll/viewport, styling, or when the user asks for Canvas rendering, 自绘表格, or to avoid DOM for the sheet.
---

# FlexSheet Canvas 渲染规范

与 [flexsheet-architecture](../flexsheet-architecture/SKILL.md) 一致：**表格主界面必须用 Canvas 2D 绘制**，不得用 DOM/React/Vue 等实现主网格。与 [flexsheet-ts-style](../flexsheet-ts-style/SKILL.md) 一致：类型与导出纪律适用于 Render 层代码。

## 硬性约束

| 允许 | 禁止 |
|------|------|
| `<canvas>` + `CanvasRenderingContext2D` 作为主渲染面 | 用 div/table 铺满单元格做「表格」主 UI |
| 编辑框等**辅助**可用单个透明/浮动 input（非网格本体） | 第三方表格组件替代自绘 |
| 样式来自数据模型或样式表，在绘制时读入 | 在 Render 层直接拼公式或改业务数据 |

## 主渲染器：`CanvasRenderer`

- **位置**：`src/render/`（或项目约定的 Render 物理目录），**不**放在 `data/`、`formula/`。
- **职责**：持有 canvas 尺寸与 DPR、协调「算可见区 → 画背景/网格/单元格 → 画覆盖层（选区、拖拽等）」、触发 `requestAnimationFrame` 或统一调度器的重绘。
- **边界**：从 Data/Feature 读取**只读**视图模型（可见行列、合并信息、冻结偏移）；不在这里改单元格值；滚动/缩放状态可与 Interaction 同步，但**绘制命令**集中在此类或其委托的 painter 中。

建议将过重逻辑拆成私有方法或 `GridPainter`、`HeaderPainter` 等，但对外入口仍以 **`CanvasRenderer`** 为统一门面（类名可保留，便于技能与代码检索一致）。

## 视口裁剪与虚拟滚动

- **视口**：用 `scrollX`/`scrollY`（或等价的 first visible row/col + offset）定义**文档坐标 → 屏幕坐标**变换；绘制前 `ctx.save()`，对内容区做 `translate(-scrollX, -scrollY)`（或等价矩阵），只绘制与 `canvas` 裁剪矩形相交的实体。
- **裁剪**：使用 `ctx.beginPath(); rect(...); clip()` 限制表体、表头、冻结区分区，避免冻结区与滚动区互相覆盖错误。
- **虚拟滚动**：**不**为屏幕外整表分配绘制对象；按「当前可见行号/列号区间 ± 缓冲行/列」遍历单元格。缓冲用于快速滚动时减少白闪，尺寸按性能调优（例如 ±1～3 行/列）。
- **尺寸与 DPR**：`canvas` 的 `width`/`height` 使用 **CSS 像素 × `devicePixelRatio`**，绘制前 `scale(dpr, dpr)` 或等价处理，保证锐化。

## 单元格绘制

- **顺序（建议）**：清屏/背景 → 网格线（可选分层）→ 单元格背景与边框 → 文本/富文本占位 → 覆盖层（选区、填充柄、冻结线）。
- **合并单元格**：绘制主格矩形一次，从属格跳过或仅参与命中测试数据，避免重复绘文字。
- **文本**：`textBaseline`、`textAlign` 与业务「水平/垂直对齐」一致；省略号或截断在测量 `measureText` 后处理。

## 样式（字体、颜色、边框、对齐、条件格式）

- **数据来源**：样式应由 **CellStyle / 条件格式规则求值结果** 合并而成（求值可在 Data 或 Feature，Render 只消费结果对象）。
- **绘制时**：将样式映射到 `ctx`：`font`、`fillStyle`、`strokeStyle`、`lineWidth`；边框按四边或外框绘制，注意与网格线分层避免双倍线宽。
- **条件格式**：不在 Canvas 里「算公式」；只根据已算好的**显示属性**（背景色、字体色、图标位等）绘制。

## 冻结窗格与滚动逻辑

- **语义**：冻结区在滚动时**相对视口固定**；仅非冻结区随 `scrollX`/`scrollY` 移动。
- **实现要点**：
  - 将表拆为 **冻结行、冻结列、交叉角、可滚动区** 等矩形，各自 clip 后绘制；滚动区应用文档平移。
  - 滚动条或手势改变的量是 **可滚动区域的 scroll**，冻结区偏移由布局常量（冻结行列数 × 行高/列宽）决定。
- **与 Interaction 协同**：命中测试（点击单元格）必须用**同一套** scroll 与冻结偏移，把屏幕坐标反算为行列，避免 Render 与 Interaction 各算一套导致错位。

## 自检清单

- [ ] 主网格是否全部为 Canvas 路径，无 DOM 单元格矩阵？
- [ ] 是否仅绘制可见区 + 缓冲，而非全表？
- [ ] `CanvasRenderer`（或等价门面）是否集中在 `render/` 且未写业务数据回写？
- [ ] 冻结与滚动是否共用同一套坐标与裁剪？
- [ ] 样式与条件格式是否以只读结果对象驱动 `ctx`？

## 延伸阅读

- 分层与禁止项： [flexsheet-architecture](../flexsheet-architecture/SKILL.md)
- 单元格级类分工（`CellRender` / `StyleRender`）： [flexsheet-cell-render](../flexsheet-cell-render/SKILL.md)
- TypeScript 与导出： [flexsheet-ts-style](../flexsheet-ts-style/SKILL.md)
