# ECMA-376 Office Open XML 规范查询

查询 ECMA-376 Office Open XML 文件格式标准，帮助解决 .xlsx/.docx/.pptx 格式的实现问题。

## 文档位置（`docs/ecma376/`）

### 规范正文（txt）

| 文件 | 内容 | 规模 |
|------|------|------|
| `Ecma-Office-Open-XML-Part-1-Fundamentals-And-Markup-Language-Reference.txt` | 核心规范：WordprocessingML / SpreadsheetML / PresentationML / DrawingML 完整元素参考 | 11MB · 248k行 |
| `Ecma-Office-Open-XML-Part-2-Open-Packaging-Conventions.txt` | OPC 包格式：ZIP 结构、Parts、Content Types、Relationships | 200KB · 4k行 |
| `Ecma-Office-Open-XML-Part-3-Markup-Compatibility-and-Extensibility.txt` | MCE：标记兼容性与扩展性命名空间 | 76KB · 1.5k行 |
| `Ecma-Office-Open-XML-Part-4-Transitional-Migration-Features.txt` | 过渡特性：VML、向后兼容元素 | 4.3MB · 83k行 |

### XML Schema（xsd）

| 目录 | 内容 |
|------|------|
| `OfficeOpenXML-XMLSchema-Strict/` | Strict 模式完整 schema（sml.xsd=SpreadsheetML, wml.xsd=WordprocessingML, pml.xsd=PresentationML, dml-*.xsd=DrawingML） |
| `OfficeOpenXML-XMLSchema-Transitional/` | Transitional 模式 schema（含 VML，Excel/WPS 实际使用） |
| `OPC-XMLSchema/` | OPC 包 schema（opc-relationships.xsd, opc-contentTypes.xsd 等） |

### 预设数据（presets/）

| 文件 | 内容 |
|------|------|
| `presetShapeDefinitions.xml` | DrawingML 所有预设形状的几何定义 |
| `presetTextWarpDefinitions.xml` | DrawingML 文字弯曲路径定义 |
| `presetTableStyles.xml` | SpreadsheetML 预设表格样式 |
| `presetCellStyles.xml` | SpreadsheetML 预设单元格样式 |

## Part 1 章节行号索引

| 章节 | 内容 | 起始行 |
|------|------|--------|
| 第12章 | SpreadsheetML 概述 | ~3146 |
| 第17章 | WordprocessingML 元素参考 | ~8230 |
| 第18章 | SpreadsheetML 元素参考 | ~75018 |
| 第19章 | PresentationML 元素参考 | ~124138 |
| 第20章 | DrawingML 框架参考 | ~134226 |
| 第22章 | Shared MLs 参考 | ~176165 |

**SpreadsheetML 常用子区域**（Part 1）：

| 内容 | 行范围 |
|------|--------|
| Workbook / Sheet 关系 | 75018–80000 |
| Worksheet / Cell | 80000–100000 |
| Styles（字体/填充/边框/格式） | 100000–110000 |
| Shared Strings | 110000–115000 |
| Charts | 115000–124000 |
| PivotCacheDefinition | ~97000 |
| PivotTableDefinition | ~104000 |

## 查询方法

### 1. 查找元素结构（用 XSD，最快）

```
Grep: pattern="pivotTableDefinition" path="docs/ecma376/OfficeOpenXML-XMLSchema-Strict/sml.xsd"
→ 直接得到该元素的 xs:element 定义，含所有属性和子元素
```

常用 XSD 文件：
- `sml.xsd` — SpreadsheetML（xlsx 核心）
- `wml.xsd` — WordprocessingML
- `pml.xsd` — PresentationML
- `dml-main.xsd` — DrawingML 基础类型
- `dml-chart.xsd` — 图表

### 2. 查找规范说明（用 txt，获取语义解释）

```
Grep: 在 Part-1 txt 中搜索元素名获取行号
Read: offset=<行号>, limit=300 读取上下文
```

### 3. 查 OPC 包结构

```
Read: Part-2 txt（整体只有 4k 行，可全读）
或 Grep OPC-XMLSchema/opc-relationships.xsd
```

### 4. 查预设值（preset 数据）

```
Read: presets/presetTableStyles.xml 获取预设样式 ID 和定义
```

## 使用策略

- **确认元素/属性是否存在** → 先查 XSD（秒级定位）
- **理解语义和约束说明** → 再查 txt 正文（找规范解释）
- **Transitional vs Strict** → Excel/WPS 实际文件用 Transitional schema，标准实现用 Strict

$ARGUMENTS
