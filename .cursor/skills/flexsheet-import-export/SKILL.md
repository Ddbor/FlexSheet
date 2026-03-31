---
name: flexsheet-import-export
description: Pure-frontend XLSX (Office Open XML) import/export for FlexSheet—ZIP packaging, pako DEFLATE, compatibility with Excel 2016+, WPS, Yongzhong. Covers multi-sheet, styles, merge, formulas, row/column sizing; no backend. Use when implementing or reviewing XLSX IO, adapters/xlsx, 导入导出, .xlsx, OOXML, ECMA-376, or Excel/WPS compatibility.
---

# FlexSheet XLSX 导入导出

**Version:** 1.0

## 目标与边界

- **格式**：标准 **XLSX**（Office Open XML，**ECMA-376**），以 **ZIP（OPC 包）** 为容器。
- **纯前端**：无后端、无服务器；在浏览器或本地用文件 API 读写。
- **兼容**：Microsoft Excel 2016+、WPS、永中；导出文件应可直接双击打开。
- **分层**：实现落在 **Adapter**（或 `src/adapters` 下的 xlsx 模块），经薄层对接 **Data/API**；不在 Render 里解析二进制。

## 依赖

| 依赖 | 用途 |
|------|------|
| **pako** | DEFLATE 压缩（ZIP 内各条目通常使用 deflate；若自建 ZIP 打包器，用 pako 压缩每个 entry 的 payload） |

若项目已有 **JSZip** 或其它符合 OPC 的 ZIP 写入器，可由其完成打包；核心约束仍是 **标准 OPC + ECMA-376 部件**，而非任意自定义二进制。

## 能力清单（须覆盖的设计点）

| 能力 | 说明 |
|------|------|
| 多工作表 | `workbook` + 多个 `sheet`，`workbook.xml` / `workbook.xml.rels` / `sheetN.xml` |
| 单元格样式 | `styles.xml`：`fonts`、`fills`、`borders`、`cellXfs` 等 |
| 字体、颜色、边框、对齐 | 映射到上述部件与 `cell` 的 `s` 属性 |
| 合并单元格 | `mergeCells` / `mergeCell` |
| 公式 | `cell` 上 `t="str"` 或共享字符串表；公式串符合 Excel 方言时注意转义 |
| 行高列宽 | `sheetFormatPr`、`row@ht`、`col@width` 等 |

导入时按逆过程解析 ZIP → XML → 写入内部 `Workbook`/单元格模型。

## 禁止项

- 不依赖服务端生成或转换 XLSX。
- 不把「裸 DEFLATE 流」当作 `.xlsx`：**`.xlsx` 必须是合法 ZIP**，含 `[Content_Types].xml`、`_rels`、至少 `xl/` 下工作簿与工作表部件等。

## 实现要点（ZIP + pako）

1. **生成**：由内存模型生成各 XML 字符串（及必要时二进制部件），按 OPC 路径组织为 **多个命名条目**。
2. **打包**：使用 **ZIP 格式**（本地文件头 + 中央目录）；每个条目 body 可用 **pako.deflate**（或等价）压缩，并写入正确的 CRC、未压缩/压缩大小。
3. **结果**：`new Blob([zipBytes], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' })`，供下载或 `File` 保存。

若仅调用 `pako.deflate(整个内容)` 而**没有** ZIP 结构与 `[Content_Types].xml`，则**不是**有效 XLSX，Excel/WPS 无法打开。

## 示例骨架（流程示意）

下列代码只表达**流程**；实际须输出完整 OPC ZIP，而非单一 deflate 流：

```ts
export async function exportXlsx(workbook: Workbook): Promise<Blob> {
  const entries = createXlsxParts(workbook); // Map<opcPath, utf8XmlOrUint8Array>
  const zipBytes = buildOpcZip(entries, { deflate: (buf) => pako.deflate(buf) });
  return new Blob([zipBytes], {
    type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
  });
}
```

- `createXlsxParts`：生成 ECMA-376 规定的部件路径与内容（与内部 `Workbook` 模型互转）。
- `buildOpcZip`：组装 ZIP（或复用成熟 ZIP 库），保证与 Excel/WPS 期望一致。

## 与架构的关系

- 数据与样式来源：**Data / Feature**；序列化规则：**Adapter**。
- 参考：[flexsheet-architecture](../flexsheet-architecture/SKILL.md)、[flexsheet-cross-framework](../flexsheet-cross-framework/SKILL.md)。

## 检查清单

- [ ] 导出是否为合法 **ZIP + OPC**？
- [ ] `[Content_Types].xml` 与 `_rels` 是否完整？
- [ ] 多表、样式、合并、公式、行列尺寸是否按 ECMA-376 映射？
- [ ] 全程无网络、无后端？
