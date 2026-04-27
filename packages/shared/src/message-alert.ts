import { attachDraggableDialogPanel } from "./dialog-draggable.js";

const ALERT_STYLE_ID = "fs-message-alert-style";

function ensureMessageAlertStyles(): void {
  if (document.getElementById(ALERT_STYLE_ID) !== null) {
    return;
  }
  const s = document.createElement("style");
  s.id = ALERT_STYLE_ID;
  s.textContent = `
.fs-message-alert-overlay {
  position: fixed;
  inset: 0;
  z-index: 10010;
  display: flex;
  align-items: center;
  justify-content: center;
  background: rgba(0, 0, 0, 0.35);
  font-family: system-ui, -apple-system, "Segoe UI", sans-serif;
}
.fs-message-alert {
  width: min(320px, calc(100vw - 32px));
  max-width: 100%;
  box-sizing: border-box;
  background: #fff;
  border-radius: 8px;
  box-shadow: 0 8px 32px rgba(0, 0, 0, 0.18);
  overflow: hidden;
}
.fs-message-alert__header {
  position: relative;
  padding: 14px 16px 8px 16px;
  border-bottom: 1px solid #edebe9;
}
.fs-message-alert__title {
  font-size: 15px;
  font-weight: 600;
  color: #323130;
  text-align: center;
}
.fs-message-alert__body {
  padding: 16px 16px 4px 16px;
  font-size: 13px;
  line-height: 1.5;
  color: #323130;
  word-wrap: break-word;
}
.fs-message-alert__footer {
  display: flex;
  justify-content: center;
  padding: 12px 16px 16px 16px;
}
.fs-message-alert__ok {
  min-width: 72px;
  padding: 7px 14px;
  font-size: 13px;
  border: none;
  border-radius: 4px;
  background: #217346;
  color: #fff;
  font-weight: 500;
  font-family: inherit;
  cursor: pointer;
}
.fs-message-alert__ok:hover {
  background: #1a5c38;
}
`;
  document.head.appendChild(s);
}

export interface ShowMessageAlertOptions {
  /** 缺省为「提示」 */
  readonly title?: string;
}

/**
 * 与 FlexSheet 提示类对话框一致的只读信息框（单「确定」），替代 `window.alert`。
 */
export function showMessageAlert(message: string, options: ShowMessageAlertOptions = {}): void {
  ensureMessageAlertStyles();
  const t = options.title;
  const titleText = t !== undefined && t.trim() !== "" ? t : "提示";
  const overlay = document.createElement("div");
  overlay.className = "fs-message-alert-overlay";
  overlay.setAttribute("role", "presentation");

  const panel = document.createElement("div");
  panel.className = "fs-message-alert";
  panel.setAttribute("role", "alertdialog");
  panel.setAttribute("aria-modal", "true");
  panel.setAttribute("aria-labelledby", "fs-message-alert-title");

  const header = document.createElement("div");
  header.className = "fs-message-alert__header";
  const titleEl = document.createElement("div");
  titleEl.id = "fs-message-alert-title";
  titleEl.className = "fs-message-alert__title";
  titleEl.textContent = titleText;
  header.appendChild(titleEl);
  const body = document.createElement("div");
  body.className = "fs-message-alert__body";
  body.textContent = message;
  const footer = document.createElement("div");
  footer.className = "fs-message-alert__footer";
  const ok = document.createElement("button");
  ok.type = "button";
  ok.className = "fs-message-alert__ok";
  ok.textContent = "确定";
  footer.appendChild(ok);

  panel.appendChild(header);
  panel.appendChild(body);
  panel.appendChild(footer);
  overlay.appendChild(panel);
  document.body.appendChild(overlay);

  attachDraggableDialogPanel(panel, header);

  const close = (): void => {
    document.removeEventListener("keydown", onKey, true);
    overlay.remove();
  };
  const onKey = (e: KeyboardEvent): void => {
    if (e.key === "Escape" || e.key === "Enter") {
      e.preventDefault();
      close();
    }
  };
  document.addEventListener("keydown", onKey, true);
  ok.addEventListener("click", () => {
    close();
  });
  overlay.addEventListener("click", (ev) => {
    if (ev.target === overlay) {
      close();
    }
  });
  requestAnimationFrame(() => {
    ok.focus();
  });
}
