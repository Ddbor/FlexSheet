---
name: flexsheet-vue-adapter
description: Defines the @flexsheet/vue adapter—thin Vue 3 bindings for framework-agnostic FlexSheet core (container mount, template ref / defineExpose to core API, props, v-model, emit events, lifecycle in script setup). Use when implementing or refactoring Vue SFC, 封装 Vue、Vue 适配器、@flexsheet/vue, v-model、defineEmits, or bridging props/events to the core API.
---

# FlexSheet Vue 适配器

**Version:** 1.0

## 目标

在 **不污染核心** 的前提下，把 FlexSheet 核心以 **Vue 3 组件** 暴露：核心仍纯 TypeScript、零 Vue 依赖；`@flexsheet/vue` 只做挂载、生命周期与 API 桥接。总原则见 [flexsheet-cross-framework](../flexsheet-cross-framework/SKILL.md)。

## 硬性规则

| 规则 | 含义 |
|------|------|
| **核心无关框架** | 核心包禁止 `import 'vue'`；适配器包可依赖 `vue`。 |
| **纯 TS + Vue 包装** | 组件用 `<script setup lang="ts">`（或 `defineComponent` + TS）；逻辑以类型明确的 props/emits/expose 为主，避免无类型 `any`。 |
| **暴露核心实例** | 用 **模板 ref** + **`defineExpose`** 暴露核心稳定 API（与 React 侧 `ref` + `useImperativeHandle` 对等），不把大量命令方法堆在 props 上。 |
| **无副作用渲染** | `setup`/渲染阶段不直接操作 Canvas、不调用核心有副作用 API；挂载、resize、同步配置进 **`onMounted` / `onUnmounted` / `watch`**（必要时 `onBeforeUnmount`）。 |
| **v-model** | 与核心支持的「值」字段对齐：用 **`defineModel`**（Vue 3.4+）或 `modelValue` + `update:modelValue`，文档中说明受控语义。 |
| **事件发射** | 用户意图与异步结果用 **`defineEmits`** 声明并 `emit`（如 `ready`、`selection-change`、`cell-edit`），命名可用 **camelCase** 或 **kebab-case** 与模板一致；内部把核心回调接到 `emit`。 |

## 组件职责

1. **容器**：模板中单一根容器（如 `div`），`ref` 指向它，供核心将 Canvas 挂入；不在 Vue 模板里主渲染表格 DOM。
2. **挂载**：在 `onMounted` 中 `mount` / `createView`，`onUnmounted`（或 `onBeforeUnmount`）对称 `destroy` / 取消订阅；若需首帧前与布局同步，可结合 `nextTick`。
3. **尺寸**：`ResizeObserver` 或 `useResizeObserver`（若项目有 composable）在 `onMounted` 注册，卸载时断开；在回调里调用核心 `resize`/`layout`，不在渲染函数里读 DOM 尺寸驱动核心。
4. **同步 props → 核心**：用 **`watch`**（或 `watchEffect`）在 props 变化时调用核心 API（如 `setData`、`setOptions`），避免在渲染期间调用。

## ref / expose 约定

- **首选**：父组件通过 **模板 ref** 拿到子组件实例，子组件 **`defineExpose({ api })`** 暴露核心对外类型（如 `FlexSheetApi`）。
- **禁止**：把每次 `setup` 重建的对象当作稳定「实例」传给核心；需用 `shallowRef`/`ref` 持有核心句柄并保持引用语义清晰。

## props / v-model / emits 约定

- **Props**：与核心配置一一对应或薄映射（Vue 习惯：`readonly`、`initialXxx` 等）。
- **v-model**：字段名与核心一致或文档说明映射；多 v-model 用 `defineModel('foo')` 等具名模型。
- **Emits**：`ready(api)`、`selection-change`、`cell-edit` 等；`defineEmits` 提供类型，与核心事件一一对齐。

## 检查清单

- [ ] 核心包仍无任何 Vue import？
- [ ] 副作用均在生命周期 / `watch`，渲染路径纯净？
- [ ] 对外能力经 `defineExpose` / props / v-model / emits，未在适配层复制公式/选区/绘制逻辑？
- [ ] 卸载时释放观察器、定时器、核心实例？

## 相关技能

- [flexsheet-cross-framework](../flexsheet-cross-framework/SKILL.md)
- [flexsheet-react-adapter](../flexsheet-react-adapter/SKILL.md)（对等能力参考）
- [flexsheet-architecture](../flexsheet-architecture/SKILL.md)
- [flexsheet-ts-style](../flexsheet-ts-style/SKILL.md)
