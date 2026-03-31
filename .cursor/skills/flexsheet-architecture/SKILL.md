---
name: flexsheet-architecture
description: Defines FlexSheet layered architecture (Base, Data, Render, Interaction, Formula, Feature, API, Adapter), src/ directory layout, and module boundaries for TypeScript + Canvas 2D. Dependencies are managed with pnpm. Use when working in FlexSheet, adding features, refactoring, reviewing structure, or when the user mentions layers, core/render/formula/adapters, or project conventions.
---

# FlexSheet 架构规范

## 技术栈

| 项 | 选择 |
|----|------|
| 语言 | TypeScript |
| 包管理 | **pnpm**（`pnpm-lock.yaml`；安装与脚本见 [flexsheet-project-init](../flexsheet-project-init/SKILL.md)） |
| 绘制 | Canvas 2D API |
| 应用构建 | Vite |
| 库打包 | Rollup |

## 分层（依赖方向：自上而下可依赖下层，避免反向耦合）

| 层 | 职责摘要 |
|----|----------|
| **Base** | 类型、常量、工具函数、与业务无关的基础设施 |
| **Data** | 单元格/表/选区等数据模型与持久化结构，不含绘制与输入 |
| **Render** | 基于 Canvas 的视图画布、布局、样式与重绘调度 |
| **Interaction** | 指针/键盘、选区与编辑手势，将用户意图转为命令或状态变更 |
| **Formula** | 公式解析、依赖图、计算与错误传播 |
| **Feature** | 具体业务能力（筛选、冻结、合并等），组合下层能力 |
| **API** | 对外稳定接口，供宿主或适配层调用，隐藏内部模块细节 |
| **Adapter** | 与外部环境对接（如导出格式、宿主生命周期），薄适配层 |

实现时可将相关代码归并到 `src/core` 等物理目录（见下），但逻辑上仍按上表分层约束依赖。

## 目录结构（约定）

```
src/
├── core/           # Base + 跨层共享内核（类型、工具、调度等）
├── data/           # Data 层模型与存储相关
├── render/         # Render：Canvas 绘制与视图状态
├── interaction/    # Interaction：输入与交互状态机（若存在）
├── formula/        # Formula：引擎与 AST
├── feature/        # Feature：业务特性模块
├── api/            # API：对外入口与门面
└── adapters/       # Adapter：外部系统适配
```

按需可增设 `src/types`、`src/constants` 等，但应保持与分层职责一致，避免在 `render` 中塞入公式核心逻辑等跨层混杂。

## 禁止项

- **DOM 作为渲染目标**：界面呈现以 Canvas 2D 为主；不要用 DOM/React/Vue 等做表格主渲染。
- **后端代码**：本仓库不包含服务端逻辑；网络或持久化若存在，仅通过 Adapter/API 边界以接口形式描述。
- **第三方 UI/表格框架依赖**：不引入替代自绘方案的表格组件库；通用工具库需克制，避免引入与架构冲突的重量级框架。

## 新增代码时的检查

1. 文件落在哪一层？是否只依赖更底层或同层抽象？
2. 绘制是否仅在 Render？数据变更是否经 Data/Formula 路径？
3. 是否违反「禁止项」？
