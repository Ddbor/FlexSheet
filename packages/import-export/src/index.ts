export { exportWorkbookToXlsxBlob, exportWorkbookToXlsxBytes } from "./export-xlsx.js";
export { importXlsxToWorkbook } from "./import-xlsx.js";
export { crc32 } from "./crc32.js";
export { unzipToMap } from "./zip-reader.js";

export function downloadXlsxBlob(blob: Blob, filename: string): void {
  if (typeof document === "undefined") {
    return;
  }
  const a = document.createElement("a");
  const url = URL.createObjectURL(blob);
  a.href = url;
  a.download = filename;
  a.rel = "noopener";
  a.click();
  URL.revokeObjectURL(url);
}
