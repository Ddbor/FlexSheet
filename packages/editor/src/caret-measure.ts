/**
 * 与 Canvas 绘制相同的 font 下，根据指针 X 求文本中插入点索引（用于双击定位光标）。
 */
export function caretOffsetFromClientX(
  text: string,
  clientX: number,
  textOriginScreenX: number,
  fontCss: string,
): number {
  const relX = clientX - textOriginScreenX;
  if (relX <= 0) {
    return 0;
  }
  const canvas = document.createElement("canvas");
  const ctx = canvas.getContext("2d");
  if (ctx === null) {
    return text.length;
  }
  ctx.font = fontCss;
  const wFull = ctx.measureText(text).width;
  if (relX >= wFull) {
    return text.length;
  }

  let lo = 0;
  let hi = text.length;
  while (lo < hi) {
    const mid = Math.floor((lo + hi + 1) / 2);
    const w = ctx.measureText(text.slice(0, mid)).width;
    if (w <= relX) {
      lo = mid;
    } else {
      hi = mid - 1;
    }
  }
  const i = lo;
  if (i >= text.length) {
    return text.length;
  }
  const wBefore = i === 0 ? 0 : ctx.measureText(text.slice(0, i)).width;
  const wAfter = ctx.measureText(text.slice(0, i + 1)).width;
  const midChar = (wBefore + wAfter) / 2;
  return relX < midChar ? i : i + 1;
}
