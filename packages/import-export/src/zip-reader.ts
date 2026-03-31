import pako from "pako";

function readU32(buf: Uint8Array, offset: number): number {
  return new DataView(buf.buffer, buf.byteOffset + offset, 4).getUint32(0, true);
}

/** 从标准 ZIP 解压为路径 → 原始字节（路径正斜杠）。 */
export function unzipToMap(buffer: ArrayBuffer): Map<string, Uint8Array> {
  const u8 = new Uint8Array(buffer);
  const dv = new DataView(u8.buffer, u8.byteOffset, u8.byteLength);
  let eocd = -1;
  const scanStart = Math.max(0, u8.length - 65557);
  for (let i = u8.length - 22; i >= scanStart; i--) {
    if (readU32(u8, i) === 0x06054b50) {
      eocd = i;
      break;
    }
  }
  if (eocd < 0) {
    throw new Error("ZIP：找不到中央目录结束标记");
  }
  const cdSize = dv.getUint32(eocd + 12, true);
  const cdOffset = dv.getUint32(eocd + 16, true);
  const out = new Map<string, Uint8Array>();
  let p = cdOffset;
  const end = cdOffset + cdSize;
  const decoder = new TextDecoder("utf-8");

  while (p < end) {
    if (readU32(u8, p) !== 0x02014b50) {
      break;
    }
    const compMethod = dv.getUint16(p + 10, true);
    const compSize = dv.getUint32(p + 20, true);
    const uncompSize = dv.getUint32(p + 24, true);
    const nameLen = dv.getUint16(p + 28, true);
    const extraLen = dv.getUint16(p + 30, true);
    const commentLen = dv.getUint16(p + 32, true);
    const localHeaderOffset = dv.getUint32(p + 42, true);
    const nameBytes = u8.subarray(p + 46, p + 46 + nameLen);
    const name = decoder.decode(nameBytes).replace(/\\/g, "/");
    p += 46 + nameLen + extraLen + commentLen;

    const lh = localHeaderOffset;
    const localNameLen = dv.getUint16(lh + 26, true);
    const localExtraLen = dv.getUint16(lh + 28, true);
    const dataStart = lh + 30 + localNameLen + localExtraLen;
    const compData = u8.subarray(dataStart, dataStart + compSize);

    let fileData: Uint8Array;
    if (compMethod === 0) {
      fileData = compData.slice();
    } else if (compMethod === 8) {
      fileData = pako.inflateRaw(compData);
    } else {
      throw new Error(`ZIP：不支持的压缩方法 ${compMethod}（${name}）`);
    }
    if (uncompSize > 0 && fileData.length !== uncompSize) {
      throw new Error(`ZIP：解压长度不符 ${name}`);
    }
    out.set(name, fileData);
  }
  return out;
}
