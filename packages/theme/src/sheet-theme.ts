/**
 * 画布与 UI 主题 token（与 XLSX 导出样式解耦，后续可序列化与切换）。
 */

export interface SheetTheme {
  name: string;
  mode: "light" | "dark";
  canvasBg: string;
  gridLineColor: string;
  headerLineColor: string;
  headerBg: string;
  headerColor: string;
  headerHoverBg: string;
  headerActiveBg: string;
  cellBg: string;
  cellColor: string;
  cellBorderColor: string;
  selectionBorderColor: string;
  selectionFillColor: string;
  activeCellBorderColor: string;
  freezeLineColor: string;
  scrollbarBg: string;
  scrollbarThumb: string;
  scrollbarThumbHover: string;
  editorBg: string;
  editorColor: string;
  editorBorder: string;
  menuBg: string;
  menuColor: string;
  menuHoverBg: string;
  menuSeparator: string;
}

export function createDefaultDarkTheme(): SheetTheme {
  return {
    name: "FlexSheet Dark",
    mode: "dark",
    canvasBg: "#1e1e1e",
    gridLineColor: "#3c3c3c",
    headerLineColor: "#505050",
    headerBg: "#2d2d2d",
    headerColor: "#e0e0e0",
    headerHoverBg: "#3a3a3a",
    headerActiveBg: "#454545",
    cellBg: "#252526",
    cellColor: "#e8e8e8",
    cellBorderColor: "#3c3c3c",
    selectionBorderColor: "#4ec9b0",
    selectionFillColor: "rgba(78, 201, 176, 0.15)",
    activeCellBorderColor: "#4ec9b0",
    freezeLineColor: "#a0a0a0",
    scrollbarBg: "#2d2d2d",
    scrollbarThumb: "#5a5a5a",
    scrollbarThumbHover: "#707070",
    editorBg: "#252526",
    editorColor: "#e8e8e8",
    editorBorder: "#4ec9b0",
    menuBg: "#2d2d2d",
    menuColor: "#e8e8e8",
    menuHoverBg: "#3a3a3a",
    menuSeparator: "#505050",
  };
}

export function createDefaultLightTheme(): SheetTheme {
  return {
    name: "FlexSheet Light",
    mode: "light",
    canvasBg: "#f5f5f5",
    gridLineColor: "#d4d4d4",
    headerLineColor: "#c0c0c0",
    headerBg: "#e8e8e8",
    headerColor: "#333333",
    headerHoverBg: "#dedede",
    headerActiveBg: "#d0d0d0",
    cellBg: "#ffffff",
    cellColor: "#1a1a1a",
    cellBorderColor: "#e0e0e0",
    selectionBorderColor: "#217346",
    selectionFillColor: "rgba(33, 115, 70, 0.12)",
    activeCellBorderColor: "#217346",
    freezeLineColor: "#808080",
    scrollbarBg: "#f0f0f0",
    scrollbarThumb: "#c4c4c4",
    scrollbarThumbHover: "#a8a8a8",
    editorBg: "#ffffff",
    editorColor: "#1a1a1a",
    editorBorder: "#217346",
    menuBg: "#ffffff",
    menuColor: "#1a1a1a",
    menuHoverBg: "#e8f5e9",
    menuSeparator: "#e0e0e0",
  };
}
