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

let fsSheetPromptBaseInjected = false;

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
