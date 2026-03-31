---
name: flexsheet-theme
description: Specifies FlexSheet multi-theme tokens (light/dark, built-in presets, custom colors), JSON-serializable SheetTheme, ThemeManager lifecycle (swap tokens + Canvas redraw, no core rebuild), system prefers-color-scheme, and independence from XLSX export. Pure TypeScript, framework-agnostic. Use when implementing or reviewing themes, dark mode, 主题/明暗/护眼/高对比, scrollbar/editor/context menu colors, or theme import/export in FlexSheet.
---

# FlexSheet 主题系统

与 [flexsheet-architecture](../flexsheet-architecture/SKILL.md) 分工：主题属于 **Base（类型与默认 token）+ Render（读 token 绘制）**；**不**进入 Data/Formula；**Adapter（XLSX）** 只读写表格数据与单元格样式，不依赖画布主题。

## 目标

- 多主题、明暗模式、用户自定义颜色；**纯前端**、**纯 TS**、**无框架依赖**。
- 切换主题：**不重建** FlexSheet / Renderer **实例**，仅更新当前 `SheetTheme` 并触发 **Canvas 重绘**（及覆盖层 DOM：编辑器、菜单、滚动条样式若用 DOM）。
- 主题 **不影响** 导出 **XLSX** 的**内容与单元格样式**（导出走数据与 OOXML 样式，与视图主题解耦）。

## 核心类型（必须一致）

对外与持久化均以 `SheetTheme` 为唯一形状；新增 token 时同步扩展接口与序列化校验。

```ts
interface SheetTheme {
  name: string;
  mode: 'light' | 'dark';
  canvasBg: string;
  gridLineColor: string;
  headerLineColor: string;
  headerBg: string;
  headerColor: string;
  headerHoverBg: string;
  headerActiveBg: string;
  cellBg: string;
  cellColor: string;
  cellBorderColor: string;
  selectionBorderColor: string;
  selectionFillColor: string;
  activeCellBorderColor: string;
  freezeLineColor: string;
  scrollbarBg: string;
  scrollbarThumb: string;
  scrollbarThumbHover: string;
  editorBg: string;
  editorColor: string;
  editorBorder: string;
  menuBg: string;
  menuColor: string;
  menuHoverBg: string;
  menuSeparator: string;
}
```

### 语义覆盖（绘制与 UI）

| 区域 | 字段 |
|------|------|
| 画布背景 | `canvasBg` |
| 网格线 | `gridLineColor` |
| 行列标题 | `headerBg`, `headerColor`, `headerLineColor`, `headerHoverBg`, `headerActiveBg` |
| 单元格默认 | `cellBg`, `cellColor`, `cellBorderColor` |
| 选区 | `selectionBorderColor`, `selectionFillColor` |
| 活动格边框 | `activeCellBorderColor` |
| 冻结分割线 | `freezeLineColor` |
| 滚动条 | `scrollbarBg`, `scrollbarThumb`, `scrollbarThumbHover` |
| 单元格编辑器 | `editorBg`, `editorColor`, `editorBorder` |
| 右键菜单 | `menuBg`, `menuColor`, `menuHoverBg`, `menuSeparator` |

## 功能规则

1. **基础模式**：`light` / `dark` 两套基底；内置主题与自定义主题均带 `mode`，便于按模式筛选或跟随系统。
2. **内置多主题（示例命名，实现时固定 id）**：清爽、商务、护眼、极黑、高对比——每种为完整 `SheetTheme`，可映射为 `themes[id]`。
3. **自定义**：在某一完整主题为底上做 `Partial<SheetTheme>` 合并，生成新 `SheetTheme`（新 `name` 区分于内置 id）。
4. **JSON**：`JSON.stringify` / `JSON.parse` 与 schema 校验（可选 `zod` 或手写守卫）；支持导出文件再导入。
5. **系统明暗**：`matchMedia('(prefers-color-scheme: dark)')` 监听变化；若开启「跟随系统」，将 `mode` 同步为 `dark`/`light` 并切换到该模式下默认或用户上次选择的主题 id。
6. **XLSX**：导入导出模块**不读取** `SheetTheme`；单元格样式以数据层/OOXML 为准，避免把画布颜色写进 xlsx。

## ThemeManager 职责（模式与示例）

- 维护 `themes: Record<string, SheetTheme>` 与 `current: SheetTheme`。
- `setTheme(id: string)`：按 **主题 id** 切换（可跨 `mode`，若 id 含模式信息则同时更新 `mode`）。
- `setMode(mode)`：切换到该 **mode** 下的默认主题或保留「当前变体」映射规则（二选一在实现中固定，并在 API 层写清）。
- `createCustomTheme(partial: Partial<SheetTheme>): SheetTheme`：`{ ...base, ...partial }`，并赋唯一 `name`。
- `getTheme(): SheetTheme` 供 `CanvasRenderer` / 编辑器 / 菜单 读取；主题变更后调用核心的 **`requestRedraw()`**（或等价）而非 `new Renderer()`。

示例骨架（逻辑参考，以仓库实现为准）：

```ts
class ThemeManager {
  private current: SheetTheme;
  private themes: Record<string, SheetTheme>;

  constructor() {
    this.themes = {
      light: defaultLightTheme,
      dark: defaultDarkTheme,
      // crisp, business, eyeCare, oled, highContrast ...
    };
    this.current = this.themes.light;
  }

  setTheme(name: string): void {
    const next = this.themes[name];
    if (next) this.current = next;
  }

  setMode(mode: 'light' | 'dark'): void {
    // 实现：切换到 mode 下默认主题，或同步 id→对应 mode 变体
    const next = mode === 'dark' ? this.themes.dark : this.themes.light;
    this.current = next;
  }

  createCustomTheme(custom: Partial<SheetTheme>): SheetTheme {
    return { ...this.current, ...custom };
  }

  getTheme(): SheetTheme {
    return this.current;
  }
}
```

## 实现要点

- **Render 层**：所有颜色从 `getTheme()` 读取，禁止在绘制函数内硬编码色值（调试除外）。
- **交互层**：编辑器、右键菜单若用 DOM，用 CSS 变量或内联 style 同步 `SheetTheme` 中对应字段。
- **测试**：至少保证 `mode` 与 `name` 序列化往返一致。若做快照或局部像素对比，须符合 [flexsheet-test](../flexsheet-test/SKILL.md)（默认避免整屏像素对比；小范围/离屏方案需与项目约定一致）。

## 自检清单

- [ ] `SheetTheme` 字段齐全且与 JSON 一致
- [ ] 切换主题仅重绘/更新样式，无重建核心实例
- [ ] XLSX 路径不依赖主题对象
- [ ] 系统 `prefers-color-scheme` 行为与产品开关一致
- [ ] 内置主题 id 与文档/常量表一致（清爽、商务、护眼、极黑、高对比）
