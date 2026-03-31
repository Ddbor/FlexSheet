---
name: flexsheet-test
description: Specifies FlexSheet testing with Vitest (happy-dom/jsdom), coverage of Workbook/worksheet data, styles, formula engine, selection/range/merge, row-column ops, theme, XLSX IO sanity, undo/redo; public API only, boundary cases, non-invasive tests under test/ or __tests__ as *.test.ts. Run tests via pnpm (e.g. pnpm test, pnpm vitest). Use when writing or reviewing unit tests, Vitest setup, render/formula/XLSX test strategy, or when the user mentions 单元测试、测试规范、flexsheet-test.
---

# FlexSheet 测试规范

**版本**: 1.0

## 框架与环境

| 项 | 约定 |
|----|------|
| 包管理 | **pnpm**（安装、锁文件、脚本见 [flexsheet-project-init](../flexsheet-project-init/SKILL.md)） |
| 测试框架 | [Vitest](https://vitest.dev/) |
| 运行环境 | 纯前端，**不**依赖真实浏览器；使用 `happy-dom` 或 `jsdom` 之一作为 DOM 全局（按需选用，项目内统一） |
| 文件名 | `*.test.ts` |
| 目录 | 独立目录 `test/` 或 `__tests__/`（通常与 `src/` 并列；可按模块分子目录），测试文件命名 `*.test.ts`，不混入业务实现文件 |

`vitest.config` 中建议：`environment: 'happy-dom'` 或 `'jsdom'`，`include` 指向上述测试目录。

## 覆盖范围（按模块）

编写或补充测试时，优先覆盖以下**公开 API**（路径随仓库实际模块名调整，如 `src/data`、`src/formula`、`src/api`）：

| 领域 | 关注点 |
|------|--------|
| 数据模型 | Workbook / Worksheet、单元格值与显示 |
| 样式 | 单元格样式读写与序列化边界 |
| 公式 | 解析与计算：`SUM`、`IF`、`VLOOKUP`、`AVERAGE` 等公开行为；非法公式、循环引用、空单元格参与计算 |
| 选区与范围 | 选区模型、范围包含/相交判断 |
| Feature | 合并单元格、行列插入删除与尺寸 |
| 主题 | `light` / `dark`（或项目等价 API）切换后 token 与依赖主题的只读行为 |
| XLSX | 导入导出**结构正确性**（工作表、单元格、合并、公式字符串等可断言部分）；不依赖后端 |
| 历史 | 撤销 / 重做对数据与选区的可观测效果 |

**渲染（Canvas）**：以**可单测的适配层与纯函数**为主（坐标、可见区域、样式键）。避免整屏像素对比，除非项目已约定快照或离屏 Canvas；**主题相关**小范围对比若采用，须与 [flexsheet-theme](../flexsheet-theme/SKILL.md) 说明一致。

## 编写规则

1. **每个对外函数或等价行为**至少一条**基础正向**用例（能表达契约即可，不堆砌重复场景）。
2. **边界**：`null`/`undefined`、空字符串、`0`、负数、极大行列数或稀疏大表、明显非法公式。
3. **隔离**：测试代码**不侵入**业务实现文件；不通过改 `src` 源码「仅为测试开洞」。优先测 `export` 的类/函数与 API 门面。
4. **异步**：仅在 API 为异步时使用 `async`/`await`；禁止无必要的 `setTimeout`、轮询等待。
5. **性能**：单测应快；避免全量构建超大工作簿除非该用例专门覆盖「大数据路径」且单独标记/分组。

## 禁止

- 冗余用例（同一行为多条等价断言）。
- 测试 **private** 方法或通过非公开符号耦合实现细节；只测**公开 API** 或可稳定契约的内部模块边界（若模块有明确 `export` 供测试）。
- 慢测试、滥用异步、无断言的占位 `it`。

## 示例结构

```typescript
import { describe, it, expect } from 'vitest'
import { Workbook } from '../src/core/workbook'

describe('Workbook', () => {
  it('should create empty workbook', () => {
    const wb = new Workbook()
    expect(wb.worksheets.length).toBe(1)
  })
})
```

导入路径按实际包入口调整（如 `@flexsheet/core` 或相对路径）。

## 新增测试时的自检

- [ ] 依赖与脚本按 [flexsheet-project-init](../flexsheet-project-init/SKILL.md) 使用 **pnpm**；Vitest + happy-dom/jsdom，`*.test.ts`
- [ ] 断言针对公开行为，而非内部实现细节
- [ ] 包含至少一条边界或错误路径（若该 API 需容错）
- [ ] 无多余异步与长时间运行
