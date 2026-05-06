export {
  DEFAULT_XLSX_EXPORT_OPTIONS,
  exportWorkbookToXlsxBlob,
  exportWorkbookToXlsxBytes,
  type XlsxExportOptions,
  type XlsxFloatingPictureExport,
} from "./export-xlsx.js";
export { decodeTextFileBytes } from "./text-file-decode.js";
export { importXlsxToWorkbook } from "./import-xlsx.js";
export {
  DEFAULT_FLEXSHEET_JSON_EXPORT_OPTIONS,
  DEFAULT_FLEXSHEET_JSON_IMPORT_OPTIONS,
  downloadJsonText,
  FLEXSHEET_JSON_FORMAT,
  FLEXSHEET_JSON_FORMAT_VERSION,
  FLEXSHEET_JSON_GENERATOR_APP,
  parseFlexSheetJson,
  serializeWorkbookToJsonDocument,
  workbookFromFlexSheetJsonDocument,
  type FlexSheetJsonDocument,
  type FlexSheetJsonExportOptions,
  type FlexSheetJsonImportOptions,
  type FlexSheetJsonParseResult,
  type FlexSheetJsonViewState,
} from "./flexsheet-json.js";
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
