---
name: flexsheet-react-adapter
description: Defines the @flexsheet/react adapter—thin React bindings for framework-agnostic FlexSheet core (container mount, ref to core API, props and callback events, effect-scoped lifecycle). Use when implementing or refactoring React components, useImperativeHandle, FlexSheet React wrapper, 封装 React、React 适配器、@flexsheet/react, or bridging props/events to the core API.
---

# FlexSheet React 适配器

**Version:** 1.0

## 目标

在 **不污染核心** 的前提下，把 FlexSheet 核心以 React 组件形式暴露：核心仍纯 TypeScript、零 React 依赖；`@flexsheet/react` 只做挂载、生命周期与 API 桥接。总原则见 [flexsheet-cross-framework](../flexsheet-cross-framework/SKILL.md)。

## 硬性规则

| 规则 | 含义 |
|------|------|
| **核心无关框架** | 核心包禁止 `import 'react'`；适配器包可依赖 `react` / `react-dom`。 |
| **使用 ref** | 用 `ref`（及必要时 `useImperativeHandle`）暴露核心实例或稳定命令面，避免把大量命令方法挂在 props 上。 |
| **无副作用渲染** | 渲染函数内不订阅全局、不直接操作 DOM/Canvas、不调用核心「有副作用」API；副作用进 `useEffect` / `useLayoutEffect`。 |
| **props / events** | 配置与数据用 **props** 传入；用户意图与异步结果用 **回调**（`onXxx`）传出，与核心事件或 API 对齐。 |

## 组件职责

1. **容器**：渲染单一容器节点（如 `div`），供核心将 Canvas 挂入；不在 React 里主渲染表格 DOM。
2. **挂载**：在 `useLayoutEffect`（需同步布局、首帧前就绪）或 `useEffect`（纯异步订阅）中 `mount` / `createView`，卸载时对称 `destroy` / 取消订阅。
3. **尺寸**：用 `ResizeObserver` 或窗口 resize，在 effect 内调用核心 `resize`/`layout`，不在 render 里读 `clientWidth` 驱动核心。
4. **同步 props → 核心**：在 effect 中根据 props 变化调用核心 API（如 `setData`、`setOptions`），避免在 render 中调用。

## ref 约定

- **首选**：`ref` 指向核心对外稳定类型（如 `FlexSheetApi`），由 `forwardRef` + `useImperativeHandle` 暴露。
- **禁止**：用 ref 存会在每次 render 重建的对象并当作「实例」传给子层；需保持引用稳定或存于 `useRef`。

## props / events 约定

- **Props**：与核心配置一一对应或做薄映射（命名用 React 习惯：`defaultXxx`、`readOnly` 等）。
- **Events**：`onReady(api)`、`onSelectionChange`、`onCellEdit` 等；内部把核心回调或事件总线接到这些 props。
- **受控 / 非受控**：若核心支持，文档中说明与 React 受控组件模式一致（值 + `onChange`）。

## 检查清单

- [ ] 核心包仍无任何 React import？
- [ ] 副作用均在 effect，render 纯净？
- [ ] 对外能力经 ref 或 props/events，未在适配层复制公式/选区/绘制逻辑？
- [ ] 卸载时释放观察器、定时器、核心实例？

## 相关技能

- [flexsheet-cross-framework](../flexsheet-cross-framework/SKILL.md)
- [flexsheet-architecture](../flexsheet-architecture/SKILL.md)
- [flexsheet-ts-style](../flexsheet-ts-style/SKILL.md)
