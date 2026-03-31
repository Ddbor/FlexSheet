---
name: flexsheet-cross-framework
description: Keeps FlexSheet core framework-agnostic (pure TypeScript) and defines adapter packages (@flexsheet/react, @flexsheet/vue) plus multi-format builds (IIFE, ESM, CJS). Monorepos use pnpm workspaces. Use when adding or refactoring framework bindings, monorepo package layout, Rollup/Vite library output, or when the user mentions 跨框架、适配器、React/Vue/Angular 包装、框架无关核心.
---

# FlexSheet 跨框架适配规范

## 目标

- **核心库**：纯 TypeScript，**零** `react` / `vue` / `@angular/*` 等框架依赖；表格绘制与状态仍在 Canvas + 架构分层（见 [flexsheet-architecture](../flexsheet-architecture/SKILL.md)）。
- **适配器包**：只做挂载、生命周期、props/inputs 与核心 API 的薄桥接，**不**把业务逻辑迁出核心或复制一份在适配层。

## 包与边界

| 包 | 职责 |
|----|------|
| **核心**（如 `@flexsheet/core` 或仓库内 `src/api` + 下层） | 数据模型、渲染、交互、公式、对外稳定 API；可发布为独立 npm 包 |
| **`@flexsheet/react`** | `FlexSheet` 组件、`ref` 暴露核心实例、`useEffect`/`useLayoutEffect` 中 mount/unmount 与 resize 观察 |
| **`@flexsheet/vue`** | 组件 + `ref`/`expose`、`<script setup>` 里 `onMounted`/`onUnmounted`、必要时 `watch` 同步 props |
| **Angular**（若存在，如 `@flexsheet/angular`） | 独立 `Component` + `OnDestroy`，模板仅容器元素；逻辑委托核心 API |

**禁止**：在核心源码中 `import` 任何框架运行时；在适配器中实现公式/选区/绘制等核心行为。

## 适配器实现要点

1. **单一数据源**：框架侧只传配置与回调；状态与命令经核心公开 API 进出。
2. **DOM 角色**：适配器可提供**容器 div**（供 Canvas 插入）；主表格仍由核心 Canvas 渲染，与架构「禁止 DOM 主渲染」一致。
3. **类型**：核心导出框架无关类型；适配器可再导出 `React.FC` 等仅存在于适配包内的类型。
4. **对等 API**：各框架适配器对外能力尽量一致（同一组 props/事件语义，命名随框架习惯：`onXxx` vs `@xxx` vs `Output`）。

## 构建产物格式

核心库与（按需）适配器发布时应提供或明确主入口：

| 格式 | 用途 |
|------|------|
| **ESM** | 现代 bundler / Vite / webpack；`"module"` / `"exports"."import"` |
| **CJS** | Node、老工具链；`"main"` 或 `"exports"."require"` |
| **IIFE** | 浏览器 `<script>` 直引、无打包场景；单文件全局暴露（如 `FlexSheet`） |

Rollup/Vite 库模式配置时：为 **core** 优先保证 ESM + CJS（及可选 IIFE）；**React/Vue 适配器**通常以 ESM 为主、CJS 为辅，与对应生态一致即可。

## Monorepo 与 pnpm

多包仓库统一使用 **pnpm workspaces**（根目录 `pnpm-workspace.yaml` + 根 `package.json`；锁文件为 **`pnpm-lock.yaml`**）。工作区内包互引用用 `workspace:` 协议；安装与构建在根执行 `pnpm install`、`pnpm -r run build` 等。发布到 **npm** registry 的包名与 `exports` 仍按各子包 `package.json` 维护，与包管理器无关。细节与初始化约定见 [flexsheet-project-init](../flexsheet-project-init/SKILL.md)。

## 新增或修改时的检查清单

- [ ] 核心变更是否仍无框架 import？
- [ ] 新 UI 行为是否落在 Render/Interaction，而非仅在某一适配器内？
- [ ] 新公开 API 是否在核心定义一次，再由各适配器转发？
- [ ] `package.json` 的 `exports` / `types` 是否与 ESM/CJS（及 IIFE）入口一致？
- [ ] 多包场景下是否用 **pnpm workspace** 与 **`pnpm-lock.yaml`** 一致维护依赖？

## 相关技能

- 分层与目录：[flexsheet-architecture](../flexsheet-architecture/SKILL.md)
- TS 风格：[flexsheet-ts-style](../flexsheet-ts-style/SKILL.md)
