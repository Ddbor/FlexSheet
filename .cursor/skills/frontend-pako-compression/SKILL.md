---
name: frontend-pako-compression
description: Guides browser-side DEFLATE/gzip compression with pako for large payloads—localStorage/IndexedDB, network transfer, and worker offload; tuning level, binary vs string, memory and quota pitfalls. Use when optimizing 大数据存储、传输体积、前端压缩、pako、DEFLATE、gzip、减少内存、localStorage 压缩、离线缓存、或评估 CompressionStream 替代方案。
---

# 前端 pako 大体积数据压缩

## 何时用 pako

- 需在 **纯前端**、**可控格式**（自定义二进制或 zlib/gzip 流）下缩小体积；或与 **ZIP/deflateRaw** 管线一致（见本仓库 [flexsheet-import-export](../flexsheet-import-export/SKILL.md)）。
- **不要**用 pako 替代标准 XLSX/ZIP 结构；OPC 包仍须正确 ZIP 布局与本项目 `zip-writer` / `zip-reader` 约定。

## API 选择（易混点）

| 场景 | 推荐 API | 说明 |
|------|-----------|------|
| 通用二进制块、自研格式 | `pako.deflate` / `pako.inflate` | zlib 封装，含头部，**非** ZIP 条目体 |
| ZIP 内单文件条目（raw） | `pako.deflateRaw` / `pako.inflateRaw` | 与 `packages/import-export` 一致 |
| HTTP 语义上的 gzip 体 | `pako.gzip` / `pako.ungzip` | 带 gzip 头/尾；与 `Content-Encoding: gzip` 对齐时用 |

输入输出类型以 **`Uint8Array`** 为主；字符串先 `TextEncoder`/`TextDecoder`，避免在中间态保留巨大 UTF-16 字符串。

## 压缩级别 `level`

- `level: 1`–`9`（默认约 6）。**更高 = 更小体积、更慢、更多 CPU**。
- 主线程大包：倾向 **4–6** 平衡 UI；极致体积且可接受耗时再升。
- 已在 **Worker** 中：可酌情 **6–9**，并测量实际帧时间与包大小。

## 存储与传输

1. **IndexedDB / Cache API**：直接存 **`Uint8Array` 或 Blob**，避免 Base64（体积约 +33%，且多一次编解码内存峰值）。
2. **localStorage**：仅字符串；若必须存入，可对 **已压缩二进制** 做 **Base64** 或 **base64url**，并预留 **配额** 与 **5MB 量级** 上限；大数据优先 **IndexedDB**。
3. **版本与魔数**：载荷前加 **1 字节版本或短魔数**，解压失败时明确降级或清空，避免静默损坏。
4. **JSON**：`JSON.stringify` 后再 `TextEncoder.encode` 再压缩；重复键多时可评估 **更紧凑序列化**（如列式、MessagePack 等）再压，但须在技能范围外单独选型。

## 性能与内存

- **峰值内存**：压缩/解压会分配输出缓冲；超大数据用 **分块**（小块 deflate 需自行设计帧协议）或 **Web Worker** 隔离主线程。
- **主线程**：超过数十 ms 的同步 `deflate` 可能卡 UI；用 **`Worker` + postMessage**（可 transfer `ArrayBuffer`）或拆任务（`requestIdleCallback` 仅适合极小块）。
- 解压后若只需流式消费，避免一次性 `inflate` 成单块超大 `Uint8Array`（除非业务必须）。

## 与原生 `CompressionStream` 对比（简述）

- **CompressionStream/DecompressionStream**：浏览器支持度需查 **caniuse**；API 为流式，适合管道与大响应。
- **pako**：**无原生流 API 时的兜底**；与现有 ZIP/raw deflate 栈一致；包体体积分层由 bundler 决定。
- 同一产品内避免 **双栈并行** 除非有兼容层；选定一种解码路径并固定魔数/版本。

## 错误与稳定性

- `inflate`/`inflateRaw` 遇损坏数据会抛错；**try/catch**，日志区分「版本不匹配」与「数据损坏」。
- 网络层若已 `Content-Encoding: gzip`，**不要**对 body 再 `ungzip` 一次（浏览器可能已解压）。

## 本仓库参考

- ZIP 条目：`packages/import-export/src/zip-writer.ts`（`deflateRaw`）、`zip-reader.ts`（`inflateRaw`）。
- 新场景（如快照、剪贴板存档）若 **不是** OPC/ZIP，用 `deflate`/`inflate` 或 gzip 即可，**勿与 ZIP raw 混用**。

## 实现检查清单

- [ ] 选对 `deflate` / `deflateRaw` / `gzip`
- [ ] 大载荷不进主线程长时间同步计算（或已 Worker）
- [ ] 持久化路径不无谓 Base64；localStorage 有配额与大小意识
- [ ] 载荷带版本号，解压失败可恢复
- [ ] 与 XLSX/ZIP 相关需求转 [flexsheet-import-export](../flexsheet-import-export/SKILL.md)
