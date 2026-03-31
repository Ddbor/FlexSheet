import pako from "pako";
import { crc32 } from "./crc32.js";

export interface ZipEntryInput {
  readonly path: string;
  readonly data: Uint8Array;
}

/**
 * 使用 pako deflateRaw 压缩各条目，生成标准 ZIP（供 .xlsx OPC）。
 */
export function buildZipArchive(entries: readonly ZipEntryInput[]): Uint8Array {
  const parts: Uint8Array[] = [];
  const central: Uint8Array[] = [];
  let offset = 0;

  const encoder = new TextEncoder();

  for (const { path, data } of entries) {
    const nameBytes = encoder.encode(path.replace(/\\/g, "/"));
    if (nameBytes.length > 0xffff) {
      throw new Error(`ZIP 路径过长: ${path}`);
    }

    const crc = crc32(data);
    let method = 0;
    let comp = data;
    if (data.length > 0) {
      const deflated = pako.deflateRaw(data, { level: 6 });
      if (deflated.length < data.length) {
        comp = deflated;
        method = 8;
      }
    } else {
      comp = new Uint8Array(0);
    }

    const uncompSize = data.length;
    const compSize = comp.length;
    const localHeaderSize = 30 + nameBytes.length;
    const local = new Uint8Array(localHeaderSize + compSize);
    const dv = new DataView(local.buffer);
    let p = 0;
    dv.setUint32(p, 0x04034b50, true);
    p += 4;
    dv.setUint16(p, 20, true);
    p += 2;
    dv.setUint16(p, 0, true);
    p += 2;
    dv.setUint16(p, method, true);
    p += 2;
    dv.setUint16(p, 0, true);
    p += 2;
    dv.setUint16(p, 0, true);
    p += 2;
    dv.setUint32(p, crc, true);
    p += 4;
    dv.setUint32(p, compSize, true);
    p += 4;
    dv.setUint32(p, uncompSize, true);
    p += 4;
    dv.setUint16(p, nameBytes.length, true);
    p += 2;
    dv.setUint16(p, 0, true);
    p += 2;
    local.set(nameBytes, p);
    p += nameBytes.length;
    local.set(comp, p);

    parts.push(local);

    const cd = new Uint8Array(46 + nameBytes.length);
    const cdv = new DataView(cd.buffer);
    let q = 0;
    cdv.setUint32(q, 0x02014b50, true);
    q += 4;
    cdv.setUint16(q, 20, true);
    q += 2;
    cdv.setUint16(q, 20, true);
    q += 2;
    cdv.setUint16(q, 0, true);
    q += 2;
    cdv.setUint16(q, method, true);
    q += 2;
    cdv.setUint16(q, 0, true);
    q += 2;
    cdv.setUint16(q, 0, true);
    q += 2;
    cdv.setUint32(q, crc, true);
    q += 4;
    cdv.setUint32(q, compSize, true);
    q += 4;
    cdv.setUint32(q, uncompSize, true);
    q += 4;
    cdv.setUint16(q, nameBytes.length, true);
    q += 2;
    cdv.setUint16(q, 0, true);
    q += 2;
    cdv.setUint16(q, 0, true);
    q += 2;
    cdv.setUint16(q, 0, true);
    q += 2;
    cdv.setUint16(q, 0, true);
    q += 2;
    cdv.setUint32(q, 0, true);
    q += 4;
    cdv.setUint32(q, offset, true);
    q += 4;
    cd.set(nameBytes, q);

    central.push(cd);
    offset += local.length;
  }

  const centralSize = central.reduce((a, b) => a + b.length, 0);
  const centralOffset = offset;
  for (const c of central) {
    parts.push(c);
  }

  const eocd = new Uint8Array(22);
  const ev = new DataView(eocd.buffer);
  ev.setUint32(0, 0x06054b50, true);
  ev.setUint16(4, 0, true);
  ev.setUint16(6, 0, true);
  ev.setUint16(8, entries.length, true);
  ev.setUint16(10, entries.length, true);
  ev.setUint32(12, centralSize, true);
  ev.setUint32(16, centralOffset, true);
  ev.setUint16(20, 0, true);
  parts.push(eocd);

  const total = parts.reduce((a, b) => a + b.length, 0);
  const out = new Uint8Array(total);
  let o = 0;
  for (const chunk of parts) {
    out.set(chunk, o);
    o += chunk.length;
  }
  return out;
}
