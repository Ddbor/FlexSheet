import { FlexSheet } from "../index.js";
import {
  applyRibbonCommandToFlexSheet,
  FlexSheetRibbon,
  ViewRibbonController,
  type HomeTabHandles,
  type RibbonCommandEvent,
} from "@flexsheet/toolbar";

const toolbar = document.getElementById("toolbar");
const root = document.getElementById("app");
const formulaBar = document.getElementById("formula-bar");
if (toolbar === null || root === null) {
  throw new Error("缺少 #toolbar 或 #app 容器");
}

toolbar.style.display = "flex";
toolbar.style.flexDirection = "column";
toolbar.style.alignItems = "stretch";
toolbar.style.padding = "0";
toolbar.style.gap = "0";

const flexSheet = new FlexSheet({
  container: root,
  formulaBar: formulaBar ?? undefined,
});

document.body.style.backgroundColor = flexSheet.getTheme().canvasBg;

let viewController: ViewRibbonController | null = null;
let homeHandles: HomeTabHandles | null = null;

const syncRibbonUndoRedo = (): void => {
  homeHandles?.syncUndoRedo(flexSheet.canUndo(), flexSheet.canRedo());
};

new FlexSheetRibbon({
  container: toolbar,
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
