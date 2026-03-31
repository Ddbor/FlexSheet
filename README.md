# FlexSheet

纯前端、框架无关的 **Canvas 2D** 电子表格内核，采用 **pnpm monorepo** 组织，强调模块化与插件化扩展。

## 特性概览

- **自绘表格**：以 Canvas 2D 为主渲染路径，不依赖第三方表格 UI 组件做栅格绘制。
- **分层架构**：Base / Data / Render / Interaction / Formula / Feature / API / Adapter，依赖自上而下，便于维护与测试。
- **插件体系**：`PluginBase`、`UIRegistry`、命令与工具栏插槽等，便于按需组合能力。
- **公式与重算**：公式解析、工作表重算、与单元格命令联动（见 `@flexsheet/formula`）。
- **主题**：明暗主题与 `ThemePlugin`（见 `@flexsheet/theme`）。
- **导入导出**：前端 XLSX（Office Open XML）相关能力（见 `@flexsheet/import-export`）。
- **演示与构建**：根目录 Vite 开发服务器；库产物支持 ESM / CJS / IIFE（见 `vite.config.ts`）。

## 技术栈

| 项 | 说明 |
|----|------|
| 语言 | TypeScript（严格模式） |
| 包管理 | pnpm 9（`pnpm-workspace.yaml`） |
| 应用与演示 | Vite 6 |
| 单元测试 | Vitest（`happy-dom`） |
| 绘制 | Canvas 2D API |

## 仓库结构（packages）

| 包名 | 职责摘要 |
|------|----------|
| `flexsheet` | 对外聚合入口：`FlexSheet`、`createDefaultWorkbook`、各插件与类型再导出 |
| `@flexsheet/core` | 工作簿/工作表、选区、命令、插件上下文等核心模型与注册 |
| `@flexsheet/renderer` | Canvas 渲染、视口与冻结布局等 |
| `@flexsheet/selection` | 选区相关逻辑 |
| `@flexsheet/scroll` | 滚动与视口联动 |
| `@flexsheet/editor` | 单元格编辑 |
| `@flexsheet/formula` | 公式解析与重算 |
| `@flexsheet/theme` | 主题与 Theme 插件 |
| `@flexsheet/import-export` | XLSX 导入导出 |
| `@flexsheet/toolbar` | Ribbon 等工具栏 UI |
| `@flexsheet/shared` | 跨包共享工具与类型 |

## 环境要求

- **Node.js**：建议当前 LTS 版本。
- **pnpm**：与根目录 `packageManager` 字段一致（`pnpm@9.15.0`），可使用 [Corepack](https://nodejs.org/api/corepack.html) 启用：`corepack enable`。

## 快速开始

```bash
# 安装依赖
pnpm install

# 本地开发（默认端口 5173，见 vite.config.ts）
pnpm dev

# 类型检查
pnpm typecheck

# 构建（各 workspace 包 + 根 Vite 库构建 + tsconfig.build）
pnpm build

# 预览构建结果
pnpm preview

# 测试
pnpm test

# 代码风格
pnpm format
pnpm lint
```

## 架构说明

逻辑分层与目录约定详见仓库内 `.cursor/skills/flexsheet-architecture/SKILL.md`（或团队内部架构文档）。核心原则：**表格主界面用 Canvas 自绘**；**无后端代码**；业务通过 API / Adapter 与外部环境对接。

## 公共 API 入口

应用或未来 `@flexsheet/react` / `@flexsheet/vue` 等适配层可从 **`flexsheet`** 包引用，例如 `FlexSheet`、`Workbook`、`CanvasRenderer`、`FormulaEnginePlugin` 等（完整列表见 `packages/flexsheet/src/index.ts`）。

---

当前版本：**0.0.1**（见各包 `package.json`）。
