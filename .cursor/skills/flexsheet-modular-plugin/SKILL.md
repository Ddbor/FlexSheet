---
name: flexsheet-modular-plugin
description: Enforces modular boundaries and a plugin-first architecture for FlexSheet—single-responsibility packages, no cross-module direct deps (interfaces/context/events), unified plugin lifecycle (install/activate/deactivate/destroy), extension points (toolbar, menus, cell render, commands), third-party registration, and OCP-friendly APIs. Use when writing or refactoring code, designing extensibility, implementing plugins, toolbars, commands, undo/redo hooks, or when the user mentions 模块化、插件化、扩展点、内核与插件、热替换、开放封闭.
---

# FlexSheet 模块化与插件化规范

与 [flexsheet-architecture](../flexsheet-architecture/SKILL.md) 分层模型配合使用：分层描述「依赖方向与职责」；本 skill 描述「如何拆模块、如何用插件承载可扩展能力」。

## 模块拆分原则

1. **单一职责**：核心模型、渲染、交互、UI 壳层、通用工具、领域事件、命令执行路径相互分离；同一文件不混合无关职责。
2. **禁止跨模块直接依赖实现细节**：模块间只通过**接口**、**注入的上下文（context）**、**领域事件**通信；不 import 其他 feature 包的具体类再 `new` 对方（测试替身除外）。
3. **公共逻辑**：抽到 `shared/`、`core/utils` 或各层明确的 `utils` 子路径，避免复制粘贴；抽离前确认无循环依赖。
4. **命名与目录**：目录名与文件名表达业务语义； barrel `index.ts` 仅 re-export 公共 API，不导出内部实现文件。

## 插件化架构要求

### 统一插件接口

每个插件实现约定形态（名称与签名可按内核定义微调，但语义保持一致）：

| 能力 | 说明 |
|------|------|
| `name` | 唯一标识，用于注册表与依赖声明 |
| `install` | 注册扩展点、绑定处理器（可幂等） |
| `activate` | 启用交互与订阅 |
| `deactivate` | 取消订阅、禁用 UI，保留注册信息以便再次激活 |
| `destroy` | 卸载资源、从注册表移除，不可再激活 |

### 依赖与通信

- 插件之间**不**硬编码互相 import；通过 **事件总线**、**命令总线** 或 **内核暴露的只读上下文** 协作。
- 支持 **启用 / 禁用 / 卸载**；**热替换**：`deactivate` → `destroy` → 加载新实现 → `install` → `activate`，由内核协调顺序。

### 内核职责边界

- **主程序 / 内核**：文档模型与生命周期、插件注册表、调度、配置与主题注入、对外稳定 API。
- **可扩展业务**（选区策略增强、编辑行为、工具栏与 Ribbon、公式函数集、导入导出格式等）：**默认以插件或插件包** 实现；内核仅保留最小契约与默认实现入口，避免把业务逻辑写死在 `core` 主路径。

## 代码规范

- **TypeScript**：先定义 `interface` / `type`，再实现；公共 API 显式导出类型；避免 `any`（与 [flexsheet-ts-style](../flexsheet-ts-style/SKILL.md) 一致）。
- **开放封闭**：新能力通过**新插件或新实现类**接入；尽量不修改内核稳定接口签名；必要变更走版本化或适配层。
- **对外 API**：`api`/facade 层清晰、稳定；内部包可使用 `internal` 路径约定或不经 barrel 导出，避免外部依赖内部文件。
- **横切能力**：**主题**、**运行时配置**、**生命周期钩子**（before/after command、sheet mount 等）由内核统一提供，插件只订阅或读取，不各自实现一套。

## 扩展能力清单

实现或评审时对照：

| 能力 | 要求 |
|------|------|
| 第三方插件 | 通过注册 API 登记，元数据含 `name`、版本、依赖（若有） |
| 菜单 / 工具栏 / 右键菜单 | 通过贡献点（contribution）注册项与命令 id，不由插件直接改 DOM 结构树硬编码 |
| 单元格渲染 | 经 Render 层约定的 hook 或 cell renderer 插件点扩展，不绕过视口与样式管线 |
| 命令系统 | 用户操作映射为命令对象，可审计、可测试；与撤销重做栈对接 |
| 撤销重做 | 命令需可逆信息或反操作；插件注册的命令遵循同一套 undo/redo 契约 |
| 事件 | 订阅与派发使用统一事件类型命名空间，避免字符串魔法值散落 |

## 编写与重构时的执行顺序

1. 判定变更属于「内核契约」还是「插件能力」— 能插件化则插件化。
2. 画依赖：只允许指向更底层、同层抽象或内核接口；出现跨 feature 依赖则改为事件或共享接口。
3. 新增扩展点时，同时补：**类型定义**、**注册入口**、**默认 no-op 或最小实现**（若适用）。
4. 自检：是否违反「插件互不相 import」「内核不堆业务」。

## 与其他 skill 的衔接

- 分层与目录：[flexsheet-architecture](../flexsheet-architecture/SKILL.md)
- 交互与命令细节：[flexsheet-interaction](../flexsheet-interaction/SKILL.md)
- 主题与配置：[flexsheet-theme](../flexsheet-theme/SKILL.md)
- 测试边界：[flexsheet-test](../flexsheet-test/SKILL.md)
