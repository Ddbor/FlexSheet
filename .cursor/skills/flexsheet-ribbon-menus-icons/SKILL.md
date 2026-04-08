---
name: flexsheet-ribbon-menus-icons
description: FlexSheet Ribbon dropdown and floating menus use content-adaptive width (no arbitrary fixed widths; viewport clamp only). Ribbon command icons stay monochrome (currentColor, no decorative multi-color SVG). Use when implementing or reviewing packages/toolbar ribbon UI, fs-dd / fs-bd-menu, align-font-border-color menus, toolbar-dropdown positioning, or icons.ts.
---

# FlexSheet Ribbon：展开菜单宽度与图标风格

## 适用范围

`packages/toolbar` 中 Ribbon 条、`.fs-dd` 下拉、带 `data-fs-floating-menu` 的浮动面板（如对齐、合并、边框、取色器等）。

## 展开菜单宽度（自适应）

**目标**：展开层宽度由**内容**决定，不因「看起来像 Excel 默认宽度」而写死上限；仅在**不超出视口**时做约束。

**做法**：

- 菜单容器使用 **`width: max-content`**（或等价：`inline-block` + 不强制 `width`），让最长标签/行决定宽度。
- **不要**为普通命令列表设置 `max-width: 200px` 这类与内容无关的上限（会截断文案或迫使换行不当）。
- **允许**保留 **`max-width: calc(100vw - 12px)`**（或同类视口安全边距），避免贴边溢出；需要时配合横向滚动或换行策略另议。
- `.fs-dd__menu` 展开后由 `syncToolbarDropdownMenuPosition` 设为 `position: fixed`；**不要**在业务代码里给菜单写死 `min-width` / `width` / `max-width` 的 inline 样式来「对齐触发器宽度」，除非产品明确要求与触发器等宽。
- 新建浮动菜单：优先复用 **`.fs-dd__menu`** 或 **`.fs-bd-menu`** 已有模式，与 `toolbar-dropdown.ts` 的开关层、 `closeAllRibbonPopups` 行为一致。
- **字体列表**等特例：可用 **`max-height` + 纵向滚动**（如 `.fs-dd__menu--font-list`），宽度仍随**最长字体名**自适应，而非固定列宽。

**反模式**：为「统一视觉」给所有下拉设相同 `min-width: 240px`；在无长文案时强行拉宽空白。

## Ribbon 图标（全部非彩色装饰）

**目标**：工具栏与 Ribbon 上的**命令图标**为 **Office 式线性/单色** 观感，随主题字色（`currentColor`）变化，**不在 SVG 路径里写死品牌色、渐变或多色填充**。

**做法**：

- 在 `packages/toolbar/src/toolbar/icons.ts`（或同类）中：优先 **`stroke="currentColor"` + `fill="none"`**，或 **`fill="currentColor"`** 配合 **`fill-opacity`** 区分层次（与现有 `svgEl` / `svgElMarkup` 一致）。
- Ribbon 上挂载的图标均从上述工厂函数产出，保证继承 **`.fs-tb-btn__icon` / `.fs-ribbon-*`** 下的 `color`。
- **例外（不算「彩色图标」）**：**取色器、主题色条、色板格子**等需要展示真实颜色的 UI，使用色块/渐变表示**数据或预览**，不替代命令图标的单色矢量风格；命令区图标本身仍为单色轮廓或单色填充形。

**反模式**：在图标 SVG 内写 `#217346`、`rgb(...)` 等固定色；使用多色 emoji 或彩色品牌 SVG 作为 Ribbon 按钮图标。

## 改动自检

- [ ] 新增/修改的展开菜单未引入与内容无关的固定 `max-width`（视口 clamp 除外）。
- [ ] 新增图标未引入路径内写死 chroma；需彩色展示时仅限色板等功能区域，且与命令图标区分。

## 相关文件（速查）

- 下拉定位与关闭：`packages/toolbar/src/toolbar/toolbar-dropdown.ts`
- 菜单与按钮样式：`packages/toolbar/src/ribbon/FlexSheetRibbon.css`（`.fs-dd__menu`、`.fs-bd-menu`、`.fs-tb-btn__icon`）
- 图标实现：`packages/toolbar/src/toolbar/icons.ts`
