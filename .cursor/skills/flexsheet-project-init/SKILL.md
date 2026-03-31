---
name: flexsheet-project-init
description: Scaffolds a new FlexSheet-style repo—directory layout under src/, package.json, strict tsconfig, Vite lib build (ESM/CJS/IIFE), index.html, and src/index.ts entry. Pure frontend, no backend. Package management uses pnpm (pnpm-lock.yaml, packageManager field). Use when initializing a FlexSheet project, 项目初始化、从零搭建、生成模板、Vite 库模式、或需要与 FlexSheet 分层目录一致的骨架。
---

# FlexSheet 项目初始化

## 何时使用

用户要求**新建/初始化** FlexSheet 风格仓库、生成骨架文件，或对齐本仓库约定的「纯 TS + Canvas 2D + Vite（库模式）」时应用本技能。

## 硬性约束

| 项 | 约定 |
|----|------|
| 项目展示名 | **FlexSheet**（`package.json` 的 `name` 可用 `flexsheet` 或 `@scope/flexsheet`） |
| 技术栈 | TypeScript + Canvas 2D + **Vite**（生产构建由 Vite 内置 **Rollup** 完成；无需单独 `rollup.config` 除非要做多包/monorepo 高级拆分） |
| 范围 | **纯前端**，无后端、无服务端代码 |
| 类型 | `strict: true`，**禁止 `any`**；核心**框架无关**（无 React/Vue 等依赖） |
| 产物 | **ESM、CJS、IIFE** 三种格式（库入口 `src/index.ts`） |
| 包管理 | **pnpm**（见下节；不混用 npm/yarn 的 lockfile） |

## 包管理（pnpm）

- **安装依赖**：`pnpm install`；开发依赖 `pnpm add -D <pkg>`，运行时依赖 `pnpm add <pkg>`。
- **执行脚本**：`pnpm run <script>` 或 `pnpm <script>`（与 `package.json` 的 `scripts` 一致）。
- **锁文件**：仅保留 **`pnpm-lock.yaml`**，勿提交 `package-lock.json` / `yarn.lock`。
- **版本锁定（推荐）**：在根 `package.json` 声明 **`packageManager`**（与 [Corepack](https://nodejs.org/api/corepack.html) 配合锁定 pnpm 主版本），例如 `"packageManager": "pnpm@9.15.0"`。
- **Monorepo**：根目录增加 **`pnpm-workspace.yaml`**（如 `packages:` 列出 `packages/*`），子包之间用 `workspace:*` 或 `workspace:` 协议引用；递归执行脚本用 `pnpm -r run <script>`。详见 [flexsheet-cross-framework](../flexsheet-cross-framework/SKILL.md)。

## 目录结构（必须生成）

在 `src/` 下创建下列子目录（可各放一个 `index.ts` 作为 barrel，或 `.gitkeep` 占位）：

```
src/
├── core/
├── render/
├── formula/
├── interaction/
├── feature/
├── api/
├── adapters/
└── utils/
```

与 [flexsheet-architecture](../flexsheet-architecture/SKILL.md) 的关系：架构中的完整分层目录含 `src/data/` 等；**初始化骨架**为最小可运行集合，故默认不强制创建 `data/`（可按项目进度增设）。增设后数据模型仍遵守分层依赖方向。

## 必须生成的文件

1. `package.json` — `type: "module"`，`scripts`：`dev`（Vite）、`build`（建议先 `tsc` 校验再 `vite build`）、`preview`；`devDependencies`：`typescript`、`vite`；可选 `@types/node`（用于 `vite.config` 中 `path`）。
2. `tsconfig.json` — 与 [flexsheet-ts-style](../flexsheet-ts-style/SKILL.md) 一致：`"strict": true`，`noImplicitAny`、`noUnusedLocals` 等按需开启；`module`/`moduleResolution` 与 Vite 推荐一致（如 `ESNext` + `bundler`）。
3. `vite.config.ts` — **`build.lib`**：`entry: 'src/index.ts'`，`name: 'FlexSheet'`（IIFE 全局名），`formats: ['es', 'cjs', 'iife']`；用 `fileName` 区分输出文件名；`rollupOptions.output` 里为 IIFE 设置 `extend: false` 等按需项。
4. `index.html` — 仅开发/演示入口，`<script type="module" src="/src/index.ts">`（或指向 `src/main.ts` 若拆分 demo）。
5. `src/index.ts` — 对外**具名导出**占位（如版本常量、后续从 `api/` 再导出）；避免默认导出与 `any`。

## `vite.config.ts` 模板（核心）

```typescript
import { defineConfig } from 'vite';

export default defineConfig({
  build: {
    lib: {
      entry: 'src/index.ts',
      name: 'FlexSheet',
      formats: ['es', 'cjs', 'iife'],
      fileName: (format): string => {
        if (format === 'es') return 'flexsheet.es.js';
        if (format === 'cjs') return 'flexsheet.cjs';
        return 'flexsheet.iife.js';
      },
    },
    rollupOptions: {
      // 核心无框架依赖时通常无需 external；若拆出 peer 依赖再列出
    },
  },
});
```

按需补充：`root`、`resolve.alias`、开发用 `server.port` 等。

## 最小模板（可直接生成）

**package.json**

```json
{
  "name": "flexsheet",
  "version": "0.0.1",
  "private": true,
  "packageManager": "pnpm@9.15.0",
  "type": "module",
  "scripts": {
    "dev": "vite",
    "build": "tsc --noEmit && vite build",
    "preview": "vite preview"
  },
  "devDependencies": {
    "@types/node": "^22.0.0",
    "typescript": "^5.6.0",
    "vite": "^6.0.0"
  }
}
```

**tsconfig.json**

```json
{
  "compilerOptions": {
    "target": "ES2022",
    "module": "ESNext",
    "moduleResolution": "bundler",
    "strict": true,
    "noImplicitAny": true,
    "noUnusedLocals": true,
    "noUnusedParameters": true,
    "skipLibCheck": true,
    "declaration": true,
    "declarationMap": true,
    "outDir": "dist",
    "rootDir": "src",
    "isolatedModules": true,
    "verbatimModuleSyntax": true
  },
  "include": ["src/**/*.ts"]
}
```

（`vite.config.ts` 由 Vite 加载；若需对其做 `tsc` 校验，可另增 `tsconfig.node.json` 仅包含该文件。）

**index.html**

```html
<!doctype html>
<html lang="zh-CN">
  <head>
    <meta charset="UTF-8" />
    <meta name="viewport" content="width=device-width, initial-scale=1.0" />
    <title>FlexSheet</title>
  </head>
  <body>
    <script type="module" src="/src/index.ts"></script>
  </body>
</html>
```

**src/index.ts**

```typescript
/** FlexSheet 库入口 — 后续从 api/ 等 barrel 聚合导出 */

export const FLEXSHEET_VERSION = '0.0.1' as const;

export function createFlexSheet(): { readonly name: 'FlexSheet' } {
  return { name: 'FlexSheet' };
}
```

## `package.json` 要点

- `exports` / `main` / `module` / `types` 与 `dist/` 下 ESM、CJS、**`.d.ts`** 路径一致（声明文件可用 `tsc --emitDeclarationOnly` 或 `vite-plugin-dts`，二选一在 README 或后续任务中说明即可）。
- 不添加 Express、数据库驱动等后端依赖。
- 与 **pnpm** 配套：锁文件为 **`pnpm-lock.yaml`**；可选 **`packageManager`** 字段锁定 pnpm 版本。

## 自检清单

- [ ] `src/` 下八个子目录已创建  
- [ ] `tsconfig` 为 `strict`，模板代码无 `any`  
- [ ] `vite build` 产出三种格式且全局名 `FlexSheet`（IIFE）  
- [ ] 无服务端源码目录或服务器脚本  
- [ ] 使用 **pnpm** 安装与运行脚本；存在 **`pnpm-lock.yaml`**（及按需 **`packageManager`** 字段）  

## 相关技能

- [flexsheet-architecture](../flexsheet-architecture/SKILL.md)  
- [flexsheet-cross-framework](../flexsheet-cross-framework/SKILL.md)  
- [flexsheet-ts-style](../flexsheet-ts-style/SKILL.md)  
