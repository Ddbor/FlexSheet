/**
 * 与右键菜单插件共用的 prompt 基础样式（删除对话框等 modifier 仍由插件注入）。
 * FlexSheet 自带对话框（如自定义排序）先调用 `ensureFsSheetPromptStyles()` 即可。
 */
export const FS_SHEET_PROMPT_BASE_STYLE_TEXT = `
.fs-sheet-prompt-overlay {
  position: fixed;
  inset: 0;
  z-index: 10002;
  display: flex;
  align-items: center;
  justify-content: center;
  background: rgba(0, 0, 0, 0.35);
  font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
}
.fs-sheet-prompt {
  width: min(268px, calc(100vw - 32px));
  max-width: 100%;
  box-sizing: border-box;
  background: #fff;
  border-radius: 8px;
  box-shadow: 0 8px 32px rgba(0, 0, 0, 0.18);
  overflow: hidden;
}
.fs-sheet-prompt__header {
  position: relative;
  padding: 14px 40px 8px 14px;
  border-bottom: 1px solid #edebe9;
}
.fs-sheet-prompt__title {
  font-size: 15px;
  font-weight: 600;
  color: #323130;
  text-align: center;
}
.fs-sheet-prompt__close {
  position: absolute;
  right: 8px;
  top: 8px;
  width: 32px;
  height: 32px;
  border: none;
  background: transparent;
  font-size: 20px;
  line-height: 1;
  color: #605e5c;
  cursor: pointer;
  border-radius: 4px;
}
.fs-sheet-prompt__close:hover {
  background: #f3f2f1;
  color: #323130;
}
.fs-sheet-prompt__body {
  padding: 16px 16px 6px 16px;
}
.fs-sheet-prompt__label {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 13px;
  color: #323130;
}
.fs-sheet-prompt__label span {
  flex-shrink: 0;
  min-width: 42px;
}
.fs-sheet-prompt__input {
  flex: 1;
  min-width: 0;
  padding: 7px 10px;
  font-size: 14px;
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  outline: none;
  box-sizing: border-box;
}
.fs-sheet-prompt__input:focus {
  border-color: #217346;
  box-shadow: 0 0 0 1px #217346 inset;
}
.fs-sheet-prompt__select {
  flex: 1;
  min-width: 0;
  padding: 7px 10px;
  font-size: 14px;
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  outline: none;
  box-sizing: border-box;
  font-family: inherit;
  background: #fff;
}
.fs-sheet-prompt__select:focus {
  border-color: #217346;
  box-shadow: 0 0 0 1px #217346 inset;
}
.fs-sheet-prompt__footer {
  display: flex;
  justify-content: space-between;
  align-items: center;
  gap: 10px;
  padding: 12px 16px 16px 16px;
}
.fs-sheet-prompt__btn {
  min-width: 72px;
  padding: 7px 14px;
  font-size: 13px;
  border-radius: 4px;
  cursor: pointer;
  font-family: inherit;
}
.fs-sheet-prompt__btn--primary {
  border: none;
  background: #217346;
  color: #fff;
  font-weight: 500;
}
.fs-sheet-prompt__btn--primary:hover {
  background: #1a5c38;
}
.fs-sheet-prompt__btn--secondary {
  border: 1px solid #c8c6c4;
  background: #fff;
  color: #323130;
}
.fs-sheet-prompt__btn--secondary:hover {
  background: #f3f2f1;
}
.fs-sheet-prompt-overlay.fs-sheet-prompt-overlay--range-pick {
  pointer-events: none;
  background: transparent;
}
.fs-sheet-prompt-overlay.fs-sheet-prompt-overlay--range-pick .fs-sheet-prompt {
  visibility: hidden;
  pointer-events: none;
  position: absolute;
  width: 0;
  height: 0;
  overflow: hidden;
}
.fs-sheet-prompt-range-pick-bar {
  display: none;
  box-sizing: border-box;
  position: fixed;
  left: 50%;
  top: 20px;
  transform: translateX(-50%);
  z-index: 10003;
  min-width: min(420px, calc(100vw - 32px));
  max-width: calc(100vw - 32px);
  padding: 10px 12px 12px;
  border-radius: 8px;
  background: #f3f2f1;
  border: 1px solid #d2d0ce;
  box-shadow: 0 4px 16px rgba(0, 0, 0, 0.12);
  pointer-events: auto;
  font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
}
.fs-sheet-prompt-overlay.fs-sheet-prompt-overlay--range-pick .fs-sheet-prompt-range-pick-bar {
  display: flex;
  flex-direction: column;
  gap: 6px;
}
.fs-sheet-prompt-range-pick-bar__title {
  font-size: 14px;
  font-weight: 600;
  color: #323130;
  text-align: center;
}
.fs-sheet-prompt-range-pick-bar__hint {
  margin: 0;
  font-size: 11px;
  color: #605e5c;
  text-align: center;
}
.fs-sheet-prompt-range-pick-bar__row {
  display: flex;
  align-items: stretch;
  gap: 6px;
}
.fs-sheet-prompt-range-pick-bar__input {
  flex: 1;
  min-width: 0;
  padding: 7px 10px;
  font-size: 13px;
  border: 1px solid #217346;
  border-radius: 4px;
  outline: none;
  background: #fff;
  color: #323130;
  box-sizing: border-box;
}
.fs-sheet-prompt-range-pick-bar__icon-wrap {
  flex-shrink: 0;
  width: 36px;
  display: flex;
  align-items: center;
  justify-content: center;
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  background: #fff;
  color: #605e5c;
}
.fs-sheet-prompt-range-pick-bar__icon-wrap svg {
  width: 18px;
  height: 18px;
}
.fs-sheet-prompt__range-pick {
  flex-shrink: 0;
  width: 36px;
  padding: 0;
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  background: #fff;
  display: flex;
  align-items: center;
  justify-content: center;
  color: #605e5c;
}
.fs-sheet-prompt__range-pick:not(:disabled) {
  cursor: pointer;
}
.fs-sheet-prompt__range-pick:not(:disabled):hover {
  background: #f3f2f1;
  color: #323130;
}
.fs-sheet-prompt__range-pick:disabled {
  cursor: not-allowed;
  opacity: 0.45;
}
.fs-sheet-prompt__range-pick svg {
  width: 18px;
  height: 18px;
}
`;

const FS_SHEET_CUSTOM_SORT_STYLE_TEXT = `
.fs-custom-sort-overlay {
  position: fixed;
  inset: 0;
  z-index: 10002;
  display: flex;
  align-items: center;
  justify-content: center;
  background: rgba(0, 0, 0, 0.35);
  font-family: system-ui, -apple-system, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
  box-sizing: border-box;
  /* 供「选项」子层绝对定位覆盖 */
  isolation: isolate;
}
.fs-custom-sort-overlay * {
  box-sizing: border-box;
}
.fs-custom-sort {
  position: relative;
  z-index: 0;
  width: min(520px, calc(100vw - 32px));
  max-height: min(560px, calc(100vh - 40px));
  display: flex;
  flex-direction: column;
  background: #ececec;
  border-radius: 10px;
  box-shadow: 0 8px 32px rgba(0, 0, 0, 0.2);
  overflow: hidden;
  color: #323130;
}
.fs-custom-sort__titlebar {
  display: flex;
  align-items: center;
  justify-content: center;
  min-height: 40px;
  padding: 10px 16px 6px 16px;
  user-select: none;
  -webkit-user-select: none;
}
.fs-custom-sort__head-title {
  text-align: center;
  font-size: 14px;
  font-weight: 600;
  color: #1b1a19;
}
.fs-custom-sort__options-scrim {
  position: absolute;
  inset: 0;
  z-index: 2;
  display: flex;
  align-items: center;
  justify-content: center;
  padding: 16px;
  background: rgba(0, 0, 0, 0.22);
  border-radius: 10px;
}
.fs-custom-sort__options-scrim[hidden] {
  display: none !important;
}
.fs-custom-sort__options-panel {
  width: min(420px, calc(100vw - 48px));
  background: #fff;
  border-radius: 8px;
  box-shadow: 0 8px 28px rgba(0, 0, 0, 0.18);
  border: 1px solid #d2d0ce;
  padding: 12px 18px 14px 18px;
  color: #323130;
  font-size: 13px;
}
.fs-custom-sort__options-drag-title {
  margin: -2px 0 12px 0;
  padding: 0 0 8px 0;
  border-bottom: 1px solid #edebe9;
  font-size: 14px;
  font-weight: 600;
  text-align: center;
  color: #1b1a19;
}
.fs-custom-sort__options-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 16px 24px;
  margin-bottom: 14px;
}
.fs-custom-sort__options-group-title {
  font-weight: 600;
  margin-bottom: 8px;
  color: #1b1a19;
}
.fs-custom-sort__options-radio {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-top: 6px;
  cursor: pointer;
}
.fs-custom-sort__options-radio input {
  accent-color: #217346;
  width: 15px;
  height: 15px;
  flex-shrink: 0;
  cursor: pointer;
}
.fs-custom-sort__options-check {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 14px;
  cursor: pointer;
}
.fs-custom-sort__options-check input {
  accent-color: #217346;
  width: 15px;
  height: 15px;
  cursor: pointer;
}
.fs-custom-sort__options-actions {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
}
.fs-custom-sort__top {
  display: flex;
  align-items: center;
  justify-content: space-between;
  gap: 12px;
  padding: 0 16px 10px 16px;
  font-size: 13px;
  color: #323130;
}
.fs-custom-sort__top-hint {
  flex: 1;
  min-width: 0;
  line-height: 1.4;
}
.fs-custom-sort__head-check {
  display: flex;
  align-items: center;
  gap: 6px;
  white-space: nowrap;
  font-size: 13px;
  cursor: pointer;
  flex-shrink: 0;
}
.fs-custom-sort__head-check input {
  width: 15px;
  height: 15px;
  accent-color: #217346;
  cursor: pointer;
}
.fs-custom-sort__table-wrap {
  margin: 0 16px 12px 16px;
  flex: 1;
  min-height: 0;
  background: #fff;
  border: 1px solid #c8c6c4;
  border-radius: 4px;
  overflow: auto;
}
.fs-custom-sort__table {
  width: 100%;
  border-collapse: separate;
  border-spacing: 0;
  table-layout: fixed;
  font-size: 12px;
}
.fs-custom-sort__th {
  background: #f3f2f1;
  color: #323130;
  font-weight: 600;
  text-align: left;
  padding: 8px 6px;
  border-bottom: 1px solid #d2d0ce;
  position: sticky;
  top: 0;
  z-index: 1;
}
.fs-custom-sort__th:not(:last-child),
.fs-custom-sort__td:not(:last-child) {
  border-right: 1px solid #e1dfdd;
}
.fs-custom-sort__th:nth-child(1) {
  width: 88px;
}
.fs-custom-sort__tr--selected {
  background: #217346;
  color: #fff;
}
.fs-custom-sort__tr--selected .fs-custom-sort__select {
  background: #186a3b;
  color: #fff;
  border-color: #0f4d2a;
}
.fs-custom-sort__tr--selected .fs-custom-sort__select option {
  background: #fff;
  color: #323130;
}
.fs-custom-sort__td--label {
  font-size: 12px;
  padding: 4px 8px;
  white-space: nowrap;
  vertical-align: middle;
}
.fs-custom-sort__td {
  padding: 2px 4px;
  vertical-align: middle;
}
.fs-custom-sort__select {
  width: 100%;
  min-width: 0;
  margin: 0;
  padding: 5px 4px 5px 6px;
  font-size: 12px;
  font-family: inherit;
  border: 1px solid #c8c6c4;
  border-radius: 3px;
  background: #fff;
  color: #323130;
  outline: none;
}
.fs-custom-sort__select:disabled,
.fs-custom-sort__tr--selected .fs-custom-sort__select:disabled {
  opacity: 0.5;
  cursor: not-allowed;
  background: #edebe9;
  color: #a19f9d;
  border-color: #c8c6c4;
}
.fs-custom-sort__tr--selected .fs-custom-sort__select:disabled {
  background: #145a32;
  border-color: #0f4d2a;
  color: #d0ddd5;
}
.fs-custom-sort__foot {
  display: flex;
  align-items: center;
  justify-content: space-between;
  padding: 4px 16px 16px 16px;
  gap: 10px;
}
.fs-custom-sort__level-btns {
  display: flex;
  border: 1px solid #a19f9d;
  border-radius: 4px;
  overflow: hidden;
  font-size: 0;
}
.fs-custom-sort__level-btn {
  min-width: 32px;
  height: 28px;
  padding: 0 8px;
  font-size: 14px;
  line-height: 1;
  border: none;
  border-right: 1px solid #a19f9d;
  background: #fff;
  color: #323130;
  cursor: pointer;
  font-family: inherit;
}
.fs-custom-sort__level-btn:last-of-type {
  border-right: none;
  min-width: 48px;
  font-size: 12px;
}
.fs-custom-sort__level-btn:disabled {
  opacity: 0.45;
  cursor: not-allowed;
}
.fs-custom-sort__level-btn:hover:not(:disabled) {
  background: #f3f2f1;
}
.fs-custom-sort__actions {
  display: flex;
  align-items: center;
  gap: 8px;
  flex-wrap: wrap;
  justify-content: flex-end;
}
.fs-custom-sort__btn {
  min-width: 72px;
  padding: 6px 14px;
  font-size: 13px;
  border-radius: 4px;
  cursor: pointer;
  font-family: inherit;
  border: 1px solid #8a8886;
  background: #fff;
  color: #323130;
}
.fs-custom-sort__btn:hover {
  background: #f3f2f1;
}
.fs-custom-sort__btn--primary {
  border: none;
  background: #217346;
  color: #fff;
  font-weight: 500;
}
.fs-custom-sort__btn--primary:hover {
  background: #1a5c38;
}
.fs-custom-sort__btn:disabled {
  opacity: 0.45;
  cursor: not-allowed;
}
`;

const FS_FIND_REPLACE_STYLE_TEXT = `
/* 无整屏变暗遮罩，便于在查找时观察表格上命中的高亮；空白处可穿透到表格。 */
.fs-fr-overlay {
  position: fixed;
  inset: 0;
  z-index: 10002;
  display: flex;
  align-items: center;
  justify-content: center;
  background: transparent;
  pointer-events: none;
  font-family: system-ui, -apple-system, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
  box-sizing: border-box;
}
.fs-fr-overlay * { box-sizing: border-box; }
.fs-fr {
  position: relative;
  pointer-events: auto;
  width: min(520px, calc(100vw - 28px));
  max-height: min(720px, calc(100vh - 28px));
  display: flex;
  flex-direction: column;
  background: #ececec;
  border-radius: 10px;
  box-shadow: 0 8px 32px rgba(0, 0, 0, 0.22);
  color: #323130;
  overflow: hidden;
}
.fs-fr__title {
  text-align: center;
  font-size: 15px;
  font-weight: 600;
  padding: 10px 16px 4px 16px;
  color: #1b1a19;
  user-select: none;
  -webkit-user-select: none;
}
.fs-fr__tabs {
  display: flex;
  align-items: stretch;
  justify-content: center;
  padding: 4px 20px 8px 20px;
  gap: 4px;
}
.fs-fr__tab {
  flex: 1 1 0;
  min-width: 0;
  max-width: 160px;
  padding: 8px 6px 9px 6px;
  font-size: 13px;
  border: 1px solid #d0cecd;
  border-bottom: none;
  border-radius: 6px 6px 0 0;
  background: #e0dfdd;
  color: #323130;
  cursor: pointer;
  font-family: inherit;
}
.fs-fr__tab[aria-selected="true"] {
  background: #fff;
  color: #1b1a19;
  font-weight: 600;
  box-shadow: 0 1px 0 #fff, 0 2px 4px rgba(0,0,0,0.04);
  border-color: #c8c6c4;
  z-index: 1;
  position: relative;
}
.fs-fr__tab:focus-visible {
  outline: 2px solid #107c41;
  outline-offset: 1px;
}
.fs-fr__body {
  flex: 1;
  min-height: 0;
  display: flex;
  flex-direction: column;
  padding: 0 12px 10px 12px;
  gap: 8px;
  overflow: auto;
}
.fs-fr__panel {
  display: flex;
  flex-direction: column;
  align-items: stretch;
  background: #e8e7e4;
  border: 1px solid #c8c6c4;
  border-radius: 6px;
  padding: 12px 12px 10px 12px;
  flex-shrink: 0;
}
.fs-fr__row {
  display: flex;
  align-items: center;
  gap: 8px;
  margin-bottom: 8px;
}
.fs-fr__row label {
  flex: 0 0 4.2em;
  text-align: right;
  font-size: 12px;
  color: #323130;
  white-space: nowrap;
}
.fs-fr__row .fs-fr__input-wrap,
.fs-fr__row--full .fs-fr__input-wrap {
  flex: 1;
  min-width: 0;
  position: relative;
  display: flex;
  align-items: stretch;
}
.fs-fr__input {
  width: 100%;
  min-width: 0;
  padding: 5px 8px;
  font-size: 12px;
  font-family: inherit;
  border: 1px solid #107c41;
  border-radius: 3px;
  background: #fff;
  color: #323130;
  outline: none;
}
.fs-fr__input--secondary {
  border: 1px solid #a19f9d;
}
.fs-fr__input:focus {
  box-shadow: 0 0 0 1px #107c41 inset;
  border-color: #107c41;
}
.fs-fr__opt-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 4px 16px;
  margin-top: 2px;
}
@media (max-width: 500px) {
  .fs-fr__opt-grid { grid-template-columns: 1fr; }
}
.fs-fr__left-col {
  display: flex;
  flex-direction: column;
  gap: 6px;
  min-width: 0;
}
.fs-fr__sel-row {
  display: flex;
  align-items: center;
  gap: 4px;
  min-width: 0;
  font-size: 12px;
  color: #323130;
}
.fs-fr__sel-row > span:first-of-type { flex: 0 0 3.2em; text-align: right; }
.fs-fr__select {
  flex: 1 1 0;
  min-width: 0;
  padding: 3px 6px 3px 5px;
  font-size: 12px;
  font-family: inherit;
  border: 1px solid #c8c6c4;
  border-radius: 2px;
  background: #fff;
  color: #323130;
  outline: none;
}
.fs-fr__select--accent { border-color: #107c41; }
.fs-fr__checks {
  display: flex;
  flex-direction: column;
  justify-content: flex-start;
  gap: 4px;
  font-size: 12px;
  color: #323130;
  padding: 0 0 0 2px;
}
.fs-fr__checks label {
  display: flex;
  align-items: center;
  gap: 6px;
  cursor: pointer;
}
.fs-fr__checks input { width: 14px; height: 14px; flex-shrink: 0; accent-color: #107c41; }
/* 与输入区同流排列，避免绝对定位压到「替换为」行 */
.fs-fr__opt-toggle {
  align-self: flex-end;
  margin-top: 2px;
  min-width: 64px;
  padding: 4px 8px;
  font-size: 12px;
  font-family: inherit;
  background: #fff;
  color: #323130;
  border: 1px solid #8a8886;
  border-radius: 4px;
  cursor: pointer;
}
.fs-fr__opt-toggle:hover { background: #f3f2f1; }
/* 「查找全部」表格式结果（与 Excel 列表区接近） */
.fs-fr__result-box {
  display: none;
  flex-direction: column;
  margin-top: 6px;
  border: 1px solid #d2d0ce;
  border-radius: 4px;
  background: #fff;
  min-height: 0;
  max-height: 220px;
  overflow: hidden;
}
.fs-fr__result-box[data-visible="1"] { display: flex; }
.fs-fr__result-scroll {
  flex: 1;
  min-height: 0;
  max-height: 180px;
  overflow: auto;
  font-size: 11px;
  color: #323130;
}
.fs-fr__result-table {
  width: 100%;
  border-collapse: collapse;
  table-layout: fixed;
  font: inherit;
}
.fs-fr__result-table thead th {
  position: sticky;
  top: 0;
  z-index: 1;
  text-align: left;
  font-weight: 600;
  padding: 4px 6px;
  background: #e8e7e4;
  color: #323130;
  border: 1px solid #c8c6c4;
  font-size: 10px;
  white-space: nowrap;
}
.fs-fr__result-table tbody td {
  padding: 3px 6px;
  border: 1px solid #d2d0ce;
  overflow: hidden;
  text-overflow: ellipsis;
  white-space: nowrap;
}
.fs-fr__result-table tbody tr {
  cursor: pointer;
  background: #fff;
}
.fs-fr__result-table tbody tr:hover { background: #f3f2f1; }
.fs-fr__result-table tbody tr.fs-fr__result-row--sel {
  background: #107c41;
  color: #fff;
}
.fs-fr__result-table tbody tr.fs-fr__result-row--sel:hover {
  background: #0e6a3a;
  color: #fff;
}
.fs-fr__result-status {
  flex-shrink: 0;
  padding: 3px 8px 4px 8px;
  font-size: 11px;
  color: #323130;
  background: #f3f2f1;
  border-top: 1px solid #d2d0ce;
}
.fs-fr__foot {
  display: flex;
  flex-wrap: wrap;
  align-items: center;
  justify-content: center;
  gap: 6px;
  padding: 2px 10px 12px 10px;
}
.fs-fr__btn {
  min-width: 64px;
  padding: 5px 8px 6px 8px;
  font-size: 12px;
  font-family: inherit;
  border-radius: 3px;
  border: 1px solid #8a8886;
  background: #fff;
  color: #1b1a19;
  cursor: pointer;
}
.fs-fr__btn:hover:not(:disabled) { background: #f3f2f1; }
.fs-fr__btn:disabled { opacity: 0.4; cursor: not-allowed; }
.fs-fr__btn--go {
  border: none;
  background: #107c41;
  color: #fff;
  font-weight: 500;
  min-width: 72px;
  padding: 5px 12px 6px 12px;
}
.fs-fr__btn--go:hover:not(:disabled) { background: #0c5c2f; }
.fs-fr[hidden] { display: none !important; }
`;

const FS_GOTO_SPECIAL_STYLE_TEXT = `
.fs-goto-special-overlay {
  position: fixed;
  inset: 0;
  z-index: 10002;
  display: flex;
  align-items: center;
  justify-content: center;
  background: rgba(0, 0, 0, 0.35);
  font-family: system-ui, -apple-system, "Segoe UI", "PingFang SC", "Microsoft YaHei", sans-serif;
  box-sizing: border-box;
}
.fs-goto-special-overlay * {
  box-sizing: border-box;
}
.fs-goto-special {
  width: min(440px, calc(100vw - 32px));
  max-height: min(520px, calc(100vh - 40px));
  display: flex;
  flex-direction: column;
  background: #ececec;
  border-radius: 10px;
  box-shadow: 0 8px 32px rgba(0, 0, 0, 0.2);
  overflow: hidden;
  color: #323130;
}
.fs-goto-special__titlebar {
  display: flex;
  align-items: center;
  justify-content: center;
  min-height: 40px;
  padding: 10px 16px 6px 16px;
  user-select: none;
  -webkit-user-select: none;
  cursor: default;
}
.fs-goto-special__title {
  text-align: center;
  font-size: 14px;
  font-weight: 600;
  color: #1b1a19;
}
.fs-goto-special__body {
  padding: 4px 18px 10px 18px;
  flex: 1;
  min-height: 0;
  overflow: auto;
}
.fs-goto-special__section-label {
  font-size: 12px;
  font-weight: 600;
  color: #605e5c;
  margin-bottom: 6px;
}
.fs-goto-special__kind-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 4px 20px;
  margin-bottom: 8px;
}
.fs-goto-special__kind-placeholder {
  min-height: 1px;
}
.fs-goto-special__radio {
  display: flex;
  align-items: center;
  gap: 8px;
  font-size: 13px;
  cursor: pointer;
  padding: 2px 0;
}
.fs-goto-special__radio input {
  width: 15px;
  height: 15px;
  flex-shrink: 0;
  accent-color: #217346;
  cursor: pointer;
}
.fs-goto-special__hr {
  height: 1px;
  background: #d2d0ce;
  margin: 6px 0 10px 0;
}
.fs-goto-special__options {
  min-height: 72px;
  font-size: 13px;
  color: #605e5c;
}
.fs-goto-special__noop {
  padding: 8px 0 4px 0;
  line-height: 1.45;
}
.fs-goto-special__subtype-grid {
  display: grid;
  grid-template-columns: 1fr 1fr;
  gap: 6px 16px;
  padding: 4px 0 2px 0;
}
.fs-goto-special__check {
  display: flex;
  align-items: center;
  gap: 8px;
  cursor: pointer;
  font-size: 13px;
  color: #323130;
}
.fs-goto-special__check input {
  width: 15px;
  height: 15px;
  accent-color: #217346;
  cursor: pointer;
}
.fs-goto-special__link-stack {
  display: flex;
  flex-direction: column;
  gap: 6px;
  padding: 4px 0 2px 0;
}
.fs-goto-special__footer {
  display: flex;
  justify-content: flex-end;
  gap: 8px;
  padding: 10px 16px 14px 16px;
  border-top: 1px solid #d2d0ce;
  background: #ececec;
}
.fs-goto-special__btn {
  min-width: 72px;
  padding: 5px 12px 6px 12px;
  font-size: 13px;
  border-radius: 4px;
  border: 1px solid #8a8886;
  background: #fff;
  color: #1b1a19;
  cursor: pointer;
  font-family: inherit;
}
.fs-goto-special__btn:hover {
  background: #f3f2f1;
}
.fs-goto-special__btn--primary {
  border-color: #107c41;
  background: #107c41;
  color: #fff;
  font-weight: 500;
}
.fs-goto-special__btn--primary:hover {
  background: #0c5c2f;
  border-color: #0c5c2f;
}
`;

let fsSheetPromptBaseInjected = false;
let fsCustomSortStylesInjected = false;
let fsFindReplaceStylesInjected = false;
let fsGotoSpecialStylesInjected = false;

export function ensureFsSheetPromptStyles(): void {
  if (fsSheetPromptBaseInjected) {
    return;
  }
  fsSheetPromptBaseInjected = true;
  const style = document.createElement("style");
  style.setAttribute("data-fs-sheet-prompt-base", "1");
  style.textContent = FS_SHEET_PROMPT_BASE_STYLE_TEXT;
  document.head.appendChild(style);
}

export function ensureCustomSortDialogStyles(): void {
  ensureFsSheetPromptStyles();
  if (fsCustomSortStylesInjected) {
    return;
  }
  fsCustomSortStylesInjected = true;
  const style = document.createElement("style");
  style.setAttribute("data-fs-custom-sort", "1");
  style.textContent = FS_SHEET_CUSTOM_SORT_STYLE_TEXT;
  document.head.appendChild(style);
}

export function ensureFindReplaceDialogStyles(): void {
  ensureFsSheetPromptStyles();
  if (fsFindReplaceStylesInjected) {
    return;
  }
  fsFindReplaceStylesInjected = true;
  const style = document.createElement("style");
  style.setAttribute("data-fs-find-replace", "1");
  style.textContent = FS_FIND_REPLACE_STYLE_TEXT;
  document.head.appendChild(style);
}

export function ensureGotoSpecialDialogStyles(): void {
  ensureFsSheetPromptStyles();
  if (fsGotoSpecialStylesInjected) {
    return;
  }
  fsGotoSpecialStylesInjected = true;
  const style = document.createElement("style");
  style.setAttribute("data-fs-goto-special", "1");
  style.textContent = FS_GOTO_SPECIAL_STYLE_TEXT;
  document.head.appendChild(style);
}
