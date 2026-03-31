---
name: flexsheet-prettier-lint
description: Runs Prettier formatting and ESLint-style lint checks via pnpm scripts or pnpm exec; prefers project scripts when present. Use when formatting code, fixing style, running lint, CI checks, or when the user mentions Prettier, ESLint, 格式化, lint, or code style.
---

# Prettier 与 Lint（FlexSheet）

## 目标

在修改 TypeScript/JavaScript 源码后，用 **Prettier** 统一格式，用 **ESLint**（或项目配置的同类工具）做静态检查；优先使用仓库已定义的 `pnpm` 脚本。

## 执行顺序

1. **读 `package.json` 的 `scripts`**，查找常见名称并优先使用：
   - 格式化：`format`、`prettier`、`prettier:write`
   - 检查：`lint`、`eslint`、`lint:check`
   - 自动修复：`lint:fix`、`eslint:fix`
2. **有对应脚本**：在仓库根目录执行 `pnpm <script>`（例如 `pnpm format`、`pnpm lint`）。
3. **无脚本但有配置文件**（如 `.prettierrc`、`eslint.config.*`、`.eslintrc.*`）：用 `pnpm exec` 调用，例如：
   - `pnpm exec prettier --write "<paths>"`
   - `pnpm exec eslint "<paths>"`（按需加 `--fix`）
4. **既无脚本也无配置**：不要臆造配置；可说明需安装 `prettier` / `eslint` 并初始化，或请用户补充团队约定后再执行。

## 变更后的习惯

对**非琐碎**的代码编辑（多文件、逻辑改动、风格混杂）：

- 在任务收尾前运行 **format**（若项目提供）再运行 **lint**。
- 若 lint 报错，优先修复代码；仅在规则明确允许或用户要求时用 `--fix`。

## 范围与路径

- 默认针对**本次改动涉及的文件或目录**；全量 `pnpm lint` / `pnpm format` 仅在用户明确要求或脚本设计为全仓库时使用。
- 尊重 `.prettierignore`、ESLint `ignorePatterns`，避免对生成物或 `dist` 等目录误格式化。

## 与 TypeScript 的关系

- `tsc --noEmit` / `pnpm typecheck` 属于**类型检查**，与 ESLint 互补；用户若同时要求「检查」，可按 `package.json` 依次执行 typecheck 与 lint。

## 可选延伸阅读

- 团队若需从零接入，可在仓库根添加 `prettier`、`eslint` 为 `devDependencies`，在 `scripts` 中暴露 `format` 与 `lint`，并提交共享配置文件。
