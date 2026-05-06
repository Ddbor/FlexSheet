export async function writeClipboardText(text: string): Promise<void> {
  try {
    if (typeof navigator !== "undefined" && navigator.clipboard?.writeText !== undefined) {
      await navigator.clipboard.writeText(text);
      return;
    }
  } catch {
    /* 降级 */
  }
  try {
    const ta = document.createElement("textarea");
    ta.value = text;
    ta.style.position = "fixed";
    ta.style.left = "-9999px";
    ta.setAttribute("readonly", "readonly");
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    document.body.removeChild(ta);
  } catch {
    /* 忽略 */
  }
}

export async function readClipboardText(): Promise<string | null> {
  try {
    if (typeof navigator !== "undefined" && navigator.clipboard?.readText !== undefined) {
      return await navigator.clipboard.readText();
    }
  } catch {
    return null;
  }
  return null;
}

/**
 * 将 data URL 图像写入系统剪贴板（需安全上下文且浏览器支持 ClipboardItem）。
 */
export async function writeClipboardImageFromDataUrl(dataUrl: string): Promise<void> {
  if (typeof ClipboardItem === "undefined" || navigator.clipboard?.write === undefined) {
    throw new Error("当前环境不支持将图片写入剪贴板");
  }
  const res = await fetch(dataUrl);
  const blob = await res.blob();
  const type =
    blob.type !== "" && blob.type.startsWith("image/") ? blob.type : "image/png";
  await navigator.clipboard.write([new ClipboardItem({ [type]: blob })]);
}
