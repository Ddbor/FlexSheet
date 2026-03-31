---
name: flexsheet-ts-style
description: Enforces FlexSheet TypeScript style—strict compiler options, interface-first shapes, no any, explicit function return types, and named exports with barrel index.ts. Use when writing or reviewing TS in FlexSheet, editing tsconfig, typings, exports, or when the user mentions strict types, interface vs type, or module public API.
---

# FlexSheet TypeScript 编码规范

与 [flexsheet-architecture](../flexsheet-architecture/SKILL.md) 分工：架构技能管分层与目录；本技能管**类型纪律与模块导出**。包管理与工具链见 [flexsheet-project-init](../flexsheet-project-init/SKILL.md)。

## 编译器（强制）

- 项目 `tsconfig` 中 **`"strict": true`**，不得关闭。
- 若拆分子配置，继承链上须保持 strict 生效；新增包/子项目同样遵守。

## 类型声明

| 规则 | 做法 |
|------|------|
| 对象形状 | **优先 `interface`**；仅在联合、交叉、映射类型等 `interface` 无法表达时用 `type`。 |
| 未知结构 | 用 **`unknown`**，收窄后再用；**禁止 `any`**（含 `@ts-ignore` 掩盖 any）。 |
| 泛型默认 | 避免默认为 `any`；必要时默认 `unknown` 或显式约束。 |

## 函数与方法

- **必须**为公开与导出函数/方法标注**显式返回类型**（`: void`、`: Promise<...>`、具体类型等）。
- 实现内部私有 `function`/方法若类型可完全推断且团队一致，可省略；**对外 API 一律显式**。

## 导出与入口

- 优先 **`export class`**、**`export function`** 等具名导出；避免默认导出（除非框架强制且单文件单组件）。
- 目录对外暴露：**统一经 `index.ts` 再导出**（barrel），内部实现文件不直接被跨目录深层引用（与分层一致时由上层 `index` 聚合）。

## 自检清单

- [ ] `strict` 未关闭；无新增 `any`
- [ ] 新对象形状用 `interface`（除非必须用 `type`）
- [ ] 导出函数/公开方法有返回类型
- [ ] 公共 API 从模块 `index.ts` 可见，导出风格统一
