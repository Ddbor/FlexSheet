---
name: flexsheet-interaction
description: Defines FlexSheet Interaction layer—pointer (click/double-click/drag), keyboard (arrows/Enter/Tab/shortcuts), SelectionModel, CellEditor, clipboard (pure frontend), drag-fill and row/column drag, context menu. Use when implementing or reviewing selection, editing, input handling, copy-paste, fill handle, or when the user mentions 选区、鼠标、键盘、编辑、复制粘贴、拖拽、右键菜单、src/interaction.
---

# FlexSheet 交互层（Interaction）

**Version:** 1.0

与 [flexsheet-architecture](../flexsheet-architecture/SKILL.md) 分工：**Interaction** 负责指针/键盘、选区与编辑手势，将用户意图转为命令或状态变更；**不**承担公式解析与主网格绘制。坐标与选区高亮绘制与 [flexsheet-cell-render](../flexsheet-cell-render/SKILL.md) 共用同一套文档坐标与视口约定。类型与导出纪律见 [flexsheet-ts-style](../flexsheet-ts-style/SKILL.md)。

## 职责与边界

| 归属 Interaction | 不归属 Interaction |
|------------------|-------------------|
| 事件路由、命中测试调度、选区/编辑/剪贴板/拖拽状态机 | 公式 AST、单元格存储结构（归 Data/Formula） |
| `SelectionModel`、`CellEditor` 控制器与命令派发 | Canvas 像素绘制（归 Render） |
| 纯前端的复制、粘贴、撤销重做命令（若存在） | 后端或网络持久化（归 Adapter/API） |

适配器仅提供容器与焦点；**核心交互逻辑保持框架无关**（见 [flexsheet-cross-framework](../flexsheet-cross-framework/SKILL.md)）。

---

## 1. 鼠标事件

须覆盖并语义稳定：

| 行为 | 要点 |
|------|------|
| **单击** | 激活单元格、移动插入点或进入编辑（与产品约定一致）；与冻结/合并格的主格对齐。 |
| **双击** | 进入单元格内联编辑（或打开富文本等，若产品支持）。 |
| **拖拽** | 区分：选区扩展、填充柄拖拽、行列分隔条拖拽、滚动条拖拽；同一指针按下后状态机互斥。 |
| **选区** | 按下起点 → 移动扩展 → 抬起确定；支持 Shift 扩展、Ctrl/Cmd 追加多区域（若产品支持）。 |

命中测试依赖 Render/Data 提供的「屏幕坐标 → 文档行列」API，避免在 Interaction 内重复布局公式。

---

## 2. 键盘事件

| 类别 | 约定 |
|------|------|
| **方向键** | 移动活动单元格或选区（与 Shift/Ctrl 组合行为与 Excel 类表格对齐）。 |
| **Enter** | 确认编辑并下移一行，或仅换行（编辑态内 vs 导航态由状态机区分）。 |
| **Tab / Shift+Tab** | 横向移动到下一/上一可编辑单元格；跳过锁定格按产品定义。 |
| **Ctrl+C / Ctrl+V / Ctrl+X** | 复制、粘贴、剪切：走剪贴板抽象（见下文），不直接操作 DOM `innerHTML` 作为主数据路径。 |
| **Ctrl+Z / Ctrl+Y**（或 Cmd） | 撤销/重做：命令栈在核心层，Interaction 只派发命令。 |

编辑态下须防止浏览器默认行为与快捷键冲突（如 Tab 失焦）；在隐藏 `textarea` 或编辑器元素上统一拦截时可文档化策略。

---

## 3. 选区模型：SelectionModel

- **单一数据源**：当前选区、活动单元格（anchor）、多区域列表（若支持）由 `SelectionModel`（或等价名称）持有；Render 只读该模型绘制覆盖层。
- **矩形规范**：用 `startRow/startCol/endRow/endCol` 表示，**包含端点**；规范化函数保证 `start ≤ end`（行列各自排序）。
- **与合并单元格**：选区显示以主格为准，键盘导航跳过从属格或合并为一块，与 Data 层合并信息一致。

**示例形状**（实现时用 `interface`，名称与项目一致即可）：

```typescript
interface Selection {
  startRow: number;
  startCol: number;
  endRow: number;
  endCol: number;
}
```

---

## 4. 单元格编辑器：CellEditor

- **单例或池化**：同一时间通常只有一个活动编辑器实例；位置/大小由当前活动单元格的文档矩形 + 视口滚动换算为屏幕 CSS 像素。
- **职责**：展示与提交字符串；Enter/Escape 行为；输入法（IME）组合态不提前提交。
- **与公式**：若单元格为公式，进入编辑时显示公式文本；提交后经 Data/Formula 路径写回并触发重算。
- **DOM**：允许单个 `input`/`textarea` 浮在 Canvas 上方；**主表仍不采用 DOM 表格**，与架构禁止项一致。

---

## 5. 复制粘贴（纯前端）

- **复制**：将选中区域序列化为 TSV（及可选 HTML 表）写入 `navigator.clipboard` 或降级到 `execCommand('copy')` + 隐藏 textarea；**不依赖服务端**。
- **粘贴**：从剪贴板读取文本，按 TSV 规则解析为矩形块，写入起始单元格并处理合并/溢出策略；可选支持从 HTML 表解析。
- **权限**：异步剪贴板 API 失败时要有降级与错误提示（或静默降级为仅键盘快捷键路径）。

---

## 6. 拖拽填充、拖拽行列

| 类型 | 要点 |
|------|------|
| **填充柄** | 选区右下角小柄按下拖拽：按产品规则扩展序列（数字递增、日期、复制或简单模式）；释放时批量写单元格并触发一次依赖更新。 |
| **拖拽行列** | 表头/行号区按下拖拽调整行高列宽，或拖拽整行/整列移动；与 `RowColSize` 单一数据源同步，结束后标记脏区。 |

避免拖拽过程中每帧全表重算；批量提交后再统一公式重算。

---

## 7. 右键菜单

- **触发**：`contextmenu` 在画布或容器上拦截，根据命中行列弹出菜单项（剪切/复制/粘贴、插入行列、删除、设置格式入口等）。
- **实现**：菜单可为少量 DOM（下拉层），或后续自绘；**业务命令**仍经 API/命令总线，不在菜单组件内写公式逻辑。

---

## 实现自检清单

- [ ] 交互状态机是否避免 Render 直接写业务数据？
- [ ] 选区与活动格是否集中在一处，重绘仅依赖模型变更？
- [ ] 剪贴板与快捷键是否在「无后端」前提下可测？
- [ ] 与冻结、合并、隐藏行列组合是否有人类可读的测试场景？
