import { FlexSheet, mountExcelBottomBar, mountGridVerticalScrollbar } from "../index.js";
import {
  applyRibbonCommandToFlexSheet,
  cellStyleToRibbonHomeFontChrome,
  cellStyleToRibbonHomeNumberFormat,
  FlexSheetRibbon,
  ViewRibbonController,
  type HomeTabHandles,
  type RibbonCommandEvent,
} from "@flexsheet/toolbar";
import { CreatePivotTableCommand } from "../pivot/pivot-table-command.js";
import { createDemoWorkbook, DEMO_FLOATING_PICTURE_SHEET_NAME } from "./demo-workbook.js";

const chromeRoot = document.getElementById("fs-sheet-chrome");
const toolbar = document.getElementById("toolbar");
const gridCanvasHost = document.getElementById("grid-canvas-host");
const gridVscrollHost = document.getElementById("grid-vscroll-host");
const bottomChrome = document.getElementById("bottom-chrome");
const formulaBar = document.getElementById("formula-bar");
if (
  chromeRoot === null ||
  toolbar === null ||
  gridCanvasHost === null ||
  gridVscrollHost === null ||
  bottomChrome === null
) {
  throw new Error(
    "缺少 #fs-sheet-chrome、#toolbar、#grid-canvas-host、#grid-vscroll-host 或 #bottom-chrome 容器",
  );
}

toolbar.style.display = "flex";
toolbar.style.flexDirection = "column";
toolbar.style.alignItems = "stretch";
toolbar.style.padding = "0";
toolbar.style.gap = "0";

const flexSheet = new FlexSheet({
  container: gridCanvasHost,
  workbook: createDemoWorkbook(),
  formulaBar: formulaBar ?? undefined,
  chromeRoot,
});

/** 演示用浮动图：与 `demo/assets/demo-penguin.png` 同源（Vite 解析 import.meta.url）。 */
const DEMO_PENGUIN_ASSET = new URL("./assets/demo-penguin.png", import.meta.url).href;

function findSheetByName(name: string) {
  const wb = flexSheet.workbook;
  for (let i = 0; i < wb.sheetCount; i++) {
    const sh = wb.getSheet(i);
    if (sh?.name === name) {
      return sh;
    }
  }
  return undefined;
}

function findSheetIndexByName(name: string): number {
  const wb = flexSheet.workbook;
  for (let i = 0; i < wb.sheetCount; i++) {
    if (wb.getSheet(i)?.name === name) {
      return i;
    }
  }
  return -1;
}

void (async (): Promise<void> => {
  try {
    const res = await fetch(DEMO_PENGUIN_ASSET);
    if (!res.ok) {
      return;
    }
    const blob = await res.blob();
    const dataUrl = await new Promise<string>((resolve, reject) => {
      const r = new FileReader();
      r.onload = (): void => resolve(r.result as string);
      r.onerror = (): void => reject(r.error ?? new Error("read failed"));
      r.readAsDataURL(blob);
    });
    const picIdx = findSheetIndexByName(DEMO_FLOATING_PICTURE_SHEET_NAME);
    if (picIdx < 0) {
      return;
    }
    flexSheet.workbook.activeSheetIndex = picIdx;
    flexSheet.selection.selectCell(8, 3);
    flexSheet.refresh();
    flexSheet.addFloatingPictureFromDataUrl(dataUrl);
  } catch {
    /* 资源缺失或解码失败时跳过，不影响表格演示 */
  }
})();

const pivotSource = findSheetByName("透视数据源");
if (pivotSource !== undefined) {
  const pivotCmd = new CreatePivotTableCommand(flexSheet.workbook, pivotSource, {
    sourceRange: { startRow: 0, endRow: 6, startCol: 0, endCol: 2 },
    hasHeaders: true,
    rowFieldCols: [0],
    columnFieldCols: [],
    filterFieldCols: [],
    valueFields: [{ col: 2, aggregate: "sum" }],
    destination: { kind: "newSheet", preferredName: "透视表示例" },
  });
  flexSheet.workspace.commands.execute(pivotCmd);
  const picIdx = findSheetIndexByName(DEMO_FLOATING_PICTURE_SHEET_NAME);
  flexSheet.workbook.activeSheetIndex = picIdx >= 0 ? picIdx : 0;
  if (picIdx >= 0) {
    flexSheet.selection.selectCell(8, 3);
    flexSheet.refresh();
  }
}

mountGridVerticalScrollbar({ container: gridVscrollHost, flexSheet });
mountExcelBottomBar({ container: bottomChrome, flexSheet });

document.body.style.backgroundColor = flexSheet.getTheme().canvasBg;

let viewController: ViewRibbonController | null = null;
let homeHandles: HomeTabHandles | null = null;

const syncRibbonUndoRedo = (): void => {
  homeHandles?.syncUndoRedo(flexSheet.canUndo(), flexSheet.canRedo());
};

const syncRibbonFontChrome = (): void => {
  if (homeHandles === null) {
    return;
  }
  const st = flexSheet.getActiveCellStyle();
  homeHandles.syncFontChrome(cellStyleToRibbonHomeFontChrome(st));
  homeHandles.syncNumberFormatChrome(cellStyleToRibbonHomeNumberFormat(st));
};

new FlexSheetRibbon({
  container: toolbar,
  backstageCoverRoot: chromeRoot,
  flexSheet,
  onHomeTabMounted: (handles) => {
    homeHandles = handles;
    syncRibbonUndoRedo();
  },
  onViewTabMounted: (handles) => {
    viewController = new ViewRibbonController({
      flexSheet,
      viewTabHandles: handles,
      onMacroRecord: (recording) => {
        window.console.info("[宏] 录制状态:", recording);
      },
    });
    viewController.syncFromRenderer();
  },
  onCommand: (ev: RibbonCommandEvent) => {
    if (viewController?.handleCommand(ev) === true) {
      return;
    }
    if (applyRibbonCommandToFlexSheet(ev, flexSheet)) {
      return;
    }
    console.info("[Ribbon]", ev.id, ev.tab, ev.payload ?? "");
  },
});

flexSheet.subscribeUndoRedo(syncRibbonUndoRedo);
flexSheet.subscribeFormattingChrome(syncRibbonFontChrome);
