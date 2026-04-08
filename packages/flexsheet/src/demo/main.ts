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
  formulaBar: formulaBar ?? undefined,
  chromeRoot,
});

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
