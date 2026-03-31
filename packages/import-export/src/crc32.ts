/** CRC-32 (PKZIP / IEEE), 无符号 32 位。 */
export function crc32(data: Uint8Array): number {
  let c = 0xffffffff;
  for (let i = 0; i < data.length; i++) {
    c ^= data[i];
    for (let k = 0; k < 8; k++) {
      c = (c >>> 1) ^ (0xedb88320 & (0 - (c & 1)));
    }
  }
  return (c ^ 0xffffffff) >>> 0;
}
