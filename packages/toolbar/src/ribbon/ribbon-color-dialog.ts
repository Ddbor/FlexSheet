/**
 * Ribbon「其他颜色」：Office 风格色盘（色相×饱和度平面 + 亮度条 + RGB + 新增/当前预览）。
 */

export interface Rgb {
  readonly r: number;
  readonly g: number;
  readonly b: number;
}

export function rgbToHsv(r: number, g: number, b: number): { h: number; s: number; v: number } {
  const rn = r / 255;
  const gn = g / 255;
  const bn = b / 255;
  const max = Math.max(rn, gn, bn);
  const min = Math.min(rn, gn, bn);
  const d = max - min;
  let h = 0;
  if (d !== 0) {
    if (max === rn) {
      h = ((gn - bn) / d + (gn < bn ? 6 : 0)) / 6;
    } else if (max === gn) {
      h = ((bn - rn) / d + 2) / 6;
    } else {
      h = ((rn - gn) / d + 4) / 6;
    }
  }
  return {
    h: h * 360,
    s: max === 0 ? 0 : d / max,
    v: max,
  };
}

export function hsvToRgb(h: number, s: number, v: number): Rgb {
  const hh = ((h % 360) + 360) % 360;
  const c = v * s;
  const x = c * (1 - Math.abs(((hh / 60) % 2) - 1));
  const m = v - c;
  let rp = 0;
  let gp = 0;
  let bp = 0;
  if (hh < 60) {
    rp = c;
    gp = x;
  } else if (hh < 120) {
    rp = x;
    gp = c;
  } else if (hh < 180) {
    gp = c;
    bp = x;
  } else if (hh < 240) {
    gp = x;
    bp = c;
  } else if (hh < 300) {
    rp = x;
    bp = c;
  } else {
    rp = c;
    bp = x;
  }
  return {
    r: Math.round((rp + m) * 255),
    g: Math.round((gp + m) * 255),
    b: Math.round((bp + m) * 255),
  };
}

function clamp255(n: number): number {
  if (!Number.isFinite(n)) {
    return 0;
  }
  return Math.max(0, Math.min(255, Math.round(n)));
}

function rgbToHex6(r: number, g: number, b: number): string {
  const rr = clamp255(r).toString(16).padStart(2, "0");
  const gg = clamp255(g).toString(16).padStart(2, "0");
  const bb = clamp255(b).toString(16).padStart(2, "0");
  return `#${rr}${gg}${bb}`.toLowerCase();
}

function parseCssHex6(hex: string): Rgb | null {
  const t = hex.trim().replace(/^#/, "");
  if (!/^[\dA-Fa-f]{6}$/.test(t)) {
    return null;
  }
  return {
    r: parseInt(t.slice(0, 2), 16),
    g: parseInt(t.slice(2, 4), 16),
    b: parseInt(t.slice(4, 6), 16),
  };
}

function ribbonIsDarkTheme(): boolean {
  const r = document.querySelector(".fs-ribbon");
  return r?.getAttribute("data-theme") === "dark";
}

/**
 * 打开与 Excel/WPS「颜色」对话框布局相近的取色器；确定返回 `#rrggbb`，取消返回 `null`。
 * @param initialCssHex6 初始 `#rrggbb`；缺省或非 6 位十六进制时按白色展示。
 */
export function showRibbonColorDialog(initialCssHex6?: string | null): Promise<string | null> {
  return new Promise((resolve) => {
    const dialogDomId = `fs-cd-${Date.now().toString(36)}-${Math.random().toString(36).slice(2, 8)}`;
    const raw = typeof initialCssHex6 === "string" ? initialCssHex6.trim() : "";
    const initialHex = /^#[\dA-Fa-f]{6}$/i.test(raw) ? raw.toLowerCase() : "#ffffff";
    const startRgb = parseCssHex6(initialHex) ?? { r: 0, g: 0, b: 0 };
    const startHsv = rgbToHsv(startRgb.r, startRgb.g, startRgb.b);

    let h = startHsv.h;
    let s = startHsv.s;
    let v = startHsv.v;
    let settled = false;
    let teardown: (() => void) | undefined;

    const backdrop = document.createElement("div");
    backdrop.className = "fs-color-dialog-backdrop";
    if (ribbonIsDarkTheme()) {
      backdrop.classList.add("fs-color-dialog-backdrop--dark");
    }
    backdrop.setAttribute("role", "presentation");

    const dialog = document.createElement("div");
    dialog.className = "fs-color-dialog";
    dialog.setAttribute("role", "dialog");
    dialog.setAttribute("aria-modal", "true");
    const titleId = `${dialogDomId}-title`;
    dialog.setAttribute("aria-labelledby", titleId);

    const header = document.createElement("div");
    header.className = "fs-color-dialog__header";
    const title = document.createElement("div");
    title.id = titleId;
    title.className = "fs-color-dialog__title";
    title.textContent = "颜色";
    const closeBtn = document.createElement("button");
    closeBtn.type = "button";
    closeBtn.className = "fs-color-dialog__close";
    closeBtn.setAttribute("aria-label", "关闭");
    closeBtn.textContent = "×";
    header.appendChild(title);
    header.appendChild(closeBtn);

    const body = document.createElement("div");
    body.className = "fs-color-dialog__body";

    const pickerRow = document.createElement("div");
    pickerRow.className = "fs-color-dialog__picker";

    const spectrumWrap = document.createElement("div");
    spectrumWrap.className = "fs-color-dialog__spectrum-wrap";
    const spectrumCv = document.createElement("canvas");
    spectrumCv.className = "fs-color-dialog__spectrum";
    spectrumCv.width = 260;
    spectrumCv.height = 200;
    spectrumWrap.appendChild(spectrumCv);
    const cross = document.createElement("div");
    cross.className = "fs-color-dialog__crosshair";
    cross.setAttribute("aria-hidden", "true");
    spectrumWrap.appendChild(cross);

    const valueWrap = document.createElement("div");
    valueWrap.className = "fs-color-dialog__value-wrap";
    const valueCv = document.createElement("canvas");
    valueCv.className = "fs-color-dialog__value";
    valueCv.width = 22;
    valueCv.height = 200;
    valueWrap.appendChild(valueCv);
    const valueThumb = document.createElement("div");
    valueThumb.className = "fs-color-dialog__value-thumb";
    valueThumb.setAttribute("aria-hidden", "true");
    valueWrap.appendChild(valueThumb);

    pickerRow.appendChild(spectrumWrap);
    pickerRow.appendChild(valueWrap);

    const controls = document.createElement("div");
    controls.className = "fs-color-dialog__controls";

    const rgbCol = document.createElement("div");
    rgbCol.className = "fs-color-dialog__rgb";

    const previewCol = document.createElement("div");
    previewCol.className = "fs-color-dialog__preview-col";
    const previewNew = document.createElement("div");
    previewNew.className = "fs-color-dialog__preview-new";
    previewNew.title = "新增";
    const previewCur = document.createElement("div");
    previewCur.className = "fs-color-dialog__preview-cur";
    previewCur.title = "当前";
    previewCur.style.backgroundColor = initialHex;
    const labNew = document.createElement("span");
    labNew.className = "fs-color-dialog__preview-label";
    labNew.textContent = "新增";
    const labCur = document.createElement("span");
    labCur.className = "fs-color-dialog__preview-label";
    labCur.textContent = "当前";
    previewCol.appendChild(labNew);
    previewCol.appendChild(previewNew);
    previewCol.appendChild(labCur);
    previewCol.appendChild(previewCur);

    const rIn = document.createElement("input");
    const gIn = document.createElement("input");
    const bIn = document.createElement("input");
    rIn.className = "fs-color-dialog__num";
    gIn.className = "fs-color-dialog__num";
    bIn.className = "fs-color-dialog__num";
    rIn.type = "number";
    gIn.type = "number";
    bIn.type = "number";
    rIn.min = "0";
    rIn.max = "255";
    gIn.min = "0";
    gIn.max = "255";
    bIn.min = "0";
    bIn.max = "255";

    const makeRgbRow = (channel: "r" | "g" | "b", label: string, input: HTMLInputElement): HTMLDivElement => {
      const row = document.createElement("div");
      row.className = "fs-color-dialog__rgb-row";
      input.id = `${dialogDomId}-${channel}`;
      const lab = document.createElement("label");
      lab.className = "fs-color-dialog__rgb-label";
      lab.textContent = label;
      lab.htmlFor = input.id;
      const cell = document.createElement("div");
      cell.className = "fs-color-dialog__spin-cell";
      const spin = document.createElement("div");
      spin.className = "fs-color-dialog__spin-btns";
      const up = document.createElement("button");
      up.type = "button";
      up.className = "fs-color-dialog__spin-btn";
      up.setAttribute("aria-label", `${label} 增加`);
      up.textContent = "▲";
      const down = document.createElement("button");
      down.type = "button";
      down.className = "fs-color-dialog__spin-btn";
      down.setAttribute("aria-label", `${label} 减少`);
      down.textContent = "▼";
      spin.appendChild(up);
      spin.appendChild(down);
      cell.appendChild(input);
      cell.appendChild(spin);
      row.appendChild(lab);
      row.appendChild(cell);
      return row;
    };

    rgbCol.appendChild(makeRgbRow("r", "红色", rIn));
    rgbCol.appendChild(makeRgbRow("g", "绿色", gIn));
    rgbCol.appendChild(makeRgbRow("b", "蓝色", bIn));

    controls.appendChild(rgbCol);
    controls.appendChild(previewCol);

    const footer = document.createElement("div");
    footer.className = "fs-color-dialog__footer";
    const okBtn = document.createElement("button");
    okBtn.type = "button";
    okBtn.className = "fs-color-dialog__ok";
    okBtn.textContent = "确定";
    const cancelBtn = document.createElement("button");
    cancelBtn.type = "button";
    cancelBtn.className = "fs-color-dialog__cancel";
    cancelBtn.textContent = "取消";
    footer.appendChild(okBtn);
    footer.appendChild(cancelBtn);

    body.appendChild(pickerRow);
    body.appendChild(controls);
    dialog.appendChild(header);
    dialog.appendChild(body);
    dialog.appendChild(footer);
    backdrop.appendChild(dialog);
    document.body.appendChild(backdrop);

    const sctx = spectrumCv.getContext("2d", { willReadFrequently: true });
    const vctx = valueCv.getContext("2d", { willReadFrequently: true });
    if (sctx === null || vctx === null) {
      backdrop.remove();
      resolve(null);
      return;
    }

    const finish = (hex: string | null): void => {
      if (settled) {
        return;
      }
      settled = true;
      teardown?.();
      resolve(hex);
    };

    const specW = spectrumCv.width;
    const specH = spectrumCv.height;
    const valW = valueCv.width;
    const valH = valueCv.height;

    const syncInputsFromHsv = (): void => {
      const { r, g, b } = hsvToRgb(h, s, v);
      rIn.value = String(clamp255(r));
      gIn.value = String(clamp255(g));
      bIn.value = String(clamp255(b));
      const hx = rgbToHex6(r, g, b);
      previewNew.style.backgroundColor = hx;
    };

    const applyRgbInputs = (): void => {
      const r = clamp255(Number(rIn.value));
      const g = clamp255(Number(gIn.value));
      const b = clamp255(Number(bIn.value));
      const nv = rgbToHsv(r, g, b);
      h = nv.h;
      s = nv.s;
      v = nv.v;
      drawAll();
      syncInputsFromHsv();
    };

    const drawSpectrum = (): void => {
      const img = sctx.createImageData(specW, specH);
      const d = img.data;
      let i = 0;
      for (let y = 0; y < specH; y++) {
        const sat = 1 - y / (specH - 1 || 1);
        for (let x = 0; x < specW; x++) {
          const hue = (x / (specW - 1 || 1)) * 360;
          const { r, g, b } = hsvToRgb(hue, sat, v);
          d[i] = r;
          d[i + 1] = g;
          d[i + 2] = b;
          d[i + 3] = 255;
          i += 4;
        }
      }
      sctx.putImageData(img, 0, 0);
    };

    const drawValueBar = (): void => {
      const img = vctx.createImageData(valW, valH);
      const d = img.data;
      let i = 0;
      for (let y = 0; y < valH; y++) {
        const vv = 1 - y / (valH - 1 || 1);
        const { r, g, b } = hsvToRgb(h, s, vv);
        for (let x = 0; x < valW; x++) {
          d[i] = r;
          d[i + 1] = g;
          d[i + 2] = b;
          d[i + 3] = 255;
          i += 4;
        }
      }
      vctx.putImageData(img, 0, 0);
    };

    const positionCrosshair = (): void => {
      const x = (h / 360) * spectrumWrap.clientWidth;
      const y = (1 - s) * spectrumWrap.clientHeight;
      cross.style.left = `${x}px`;
      cross.style.top = `${y}px`;
    };

    const positionValueThumb = (): void => {
      const y = (1 - v) * valueWrap.clientHeight;
      valueThumb.style.top = `${y}px`;
    };

    const drawAll = (): void => {
      drawSpectrum();
      drawValueBar();
      positionCrosshair();
      positionValueThumb();
    };

    const readHsFromClient = (clientX: number, clientY: number): void => {
      const r = spectrumWrap.getBoundingClientRect();
      const x = Math.max(0, Math.min(r.width, clientX - r.left));
      const y = Math.max(0, Math.min(r.height, clientY - r.top));
      h = (x / (r.width || 1)) * 360;
      s = 1 - y / (r.height || 1);
      drawAll();
      syncInputsFromHsv();
    };

    const readVFromClient = (clientY: number): void => {
      const r = valueWrap.getBoundingClientRect();
      const y = Math.max(0, Math.min(r.height, clientY - r.top));
      v = 1 - y / (r.height || 1);
      drawAll();
      syncInputsFromHsv();
    };

    let draggingSpec = false;
    let draggingVal = false;

    spectrumWrap.addEventListener("pointerdown", (ev) => {
      if (ev.button !== 0) {
        return;
      }
      ev.preventDefault();
      draggingSpec = true;
      spectrumWrap.setPointerCapture(ev.pointerId);
      readHsFromClient(ev.clientX, ev.clientY);
    });
    spectrumWrap.addEventListener("pointermove", (ev) => {
      if (!draggingSpec) {
        return;
      }
      readHsFromClient(ev.clientX, ev.clientY);
    });
    spectrumWrap.addEventListener("pointerup", (ev) => {
      if (draggingSpec) {
        draggingSpec = false;
        try {
          spectrumWrap.releasePointerCapture(ev.pointerId);
        } catch {
          /* not capturing */
        }
      }
    });
    spectrumWrap.addEventListener("pointercancel", () => {
      draggingSpec = false;
    });

    valueWrap.addEventListener("pointerdown", (ev) => {
      if (ev.button !== 0) {
        return;
      }
      ev.preventDefault();
      draggingVal = true;
      valueWrap.setPointerCapture(ev.pointerId);
      readVFromClient(ev.clientY);
    });
    valueWrap.addEventListener("pointermove", (ev) => {
      if (!draggingVal) {
        return;
      }
      readVFromClient(ev.clientY);
    });
    valueWrap.addEventListener("pointerup", (ev) => {
      if (draggingVal) {
        draggingVal = false;
        try {
          valueWrap.releasePointerCapture(ev.pointerId);
        } catch {
          /* not capturing */
        }
      }
    });
    valueWrap.addEventListener("pointercancel", () => {
      draggingVal = false;
    });

    const wireSpinner = (input: HTMLInputElement, delta: number): void => {
      const n = clamp255(Number(input.value) + delta);
      input.value = String(n);
      applyRgbInputs();
    };

    const bindRgbSpinners = (input: HTMLInputElement): void => {
      const cell = input.closest(".fs-color-dialog__spin-cell");
      if (cell === null) {
        return;
      }
      const btns = cell.querySelectorAll(".fs-color-dialog__spin-btn");
      const upB = btns[0];
      const downB = btns[1];
      if (!(upB instanceof HTMLButtonElement) || !(downB instanceof HTMLButtonElement)) {
        return;
      }
      upB.addEventListener("click", () => wireSpinner(input, 1));
      downB.addEventListener("click", () => wireSpinner(input, -1));
    };
    bindRgbSpinners(rIn);
    bindRgbSpinners(gIn);
    bindRgbSpinners(bIn);

    rIn.addEventListener("change", () => applyRgbInputs());
    gIn.addEventListener("change", () => applyRgbInputs());
    bIn.addEventListener("change", () => applyRgbInputs());
    rIn.addEventListener("input", () => applyRgbInputs());
    gIn.addEventListener("input", () => applyRgbInputs());
    bIn.addEventListener("input", () => applyRgbInputs());

    okBtn.addEventListener("click", () => {
      const { r, g, b } = hsvToRgb(h, s, v);
      finish(rgbToHex6(r, g, b));
    });
    cancelBtn.addEventListener("click", () => finish(null));
    closeBtn.addEventListener("click", () => finish(null));
    backdrop.addEventListener("click", (ev) => {
      if (ev.target === backdrop) {
        finish(null);
      }
    });

    const onKey = (ev: KeyboardEvent): void => {
      if (ev.key === "Escape") {
        ev.preventDefault();
        finish(null);
      }
    };
    document.addEventListener("keydown", onKey, true);

    const ro = new ResizeObserver(() => {
      drawAll();
    });
    ro.observe(spectrumWrap);
    ro.observe(valueWrap);

    teardown = (): void => {
      document.removeEventListener("keydown", onKey, true);
      ro.disconnect();
      backdrop.remove();
    };

    syncInputsFromHsv();
    queueMicrotask(() => {
      drawAll();
      okBtn.focus();
    });
  });
}
