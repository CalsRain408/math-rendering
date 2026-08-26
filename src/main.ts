import "./style.css";
import {
  insertFormula,
  isInsideOffice,
  isWord,
  officeHostName,
  recoverSelectedFormula,
  supportsFormulaRoundTrip,
  type FormulaReplacementTarget,
  type InsertPreference,
} from "./office";
import {
  MATHTYPE_SIZE_RATIOS,
  renderFormula,
  svgToPngBase64,
  type RenderOptions,
  type RenderedFormula,
} from "./renderer";

const SAMPLE = String.raw`\int_{-\infty}^{\infty} e^{-x^2}\,dx = \sqrt{\pi}`;

document.querySelector<HTMLDivElement>("#app")!.innerHTML = `
  <main class="shell">
    <header class="brand">
      <div class="brand-mark" aria-hidden="true">∑</div>
      <div>
        <h1>Formula Canvas</h1>
        <p>LaTeX → 无损 Office 公式</p>
      </div>
      <span id="hostBadge" class="host-badge">正在连接…</span>
    </header>

    <section class="editor-card">
      <div class="section-head">
        <label for="formula">LaTeX / KaTeX</label>
        <div class="editor-actions">
          <button id="editSelectedButton" class="text-button" type="button" hidden>编辑选中公式</button>
          <button id="sampleButton" class="text-button" type="button">示例</button>
          <button id="clearButton" class="text-button" type="button" title="Control + Shift + K">
            清空 <kbd>⌃⇧K</kbd>
          </button>
        </div>
      </div>
      <textarea id="formula" spellcheck="false" aria-label="公式源码" placeholder="例如：\\frac{-b \\pm \\sqrt{b^2-4ac}}{2a}"></textarea>
      <div id="error" class="message error" role="alert" hidden></div>
    </section>

    <section class="preview-card">
      <div class="section-head">
        <span>预览</span>
        <span id="dimensions" class="dimensions"></span>
      </div>
      <div id="preview" class="preview" aria-live="polite"></div>
    </section>

    <section class="settings" aria-label="渲染设置">
      <label class="setting">
        <span>完整字号（Full）</span>
        <span class="number-field"><input id="fontSize" type="number" list="fontSizePresets" min="6" max="96" step="0.5" value="12" /><i>pt</i></span>
        <datalist id="fontSizePresets">
          <option value="9"></option><option value="10"></option><option value="10.5"></option>
          <option value="11"></option><option value="12" label="MathType 默认"></option>
          <option value="14"></option><option value="16"></option><option value="18"></option>
          <option value="24"></option>
        </datalist>
        <small id="sizeHint" class="setting-hint"></small>
      </label>
      <label class="setting">
        <span>颜色</span>
        <span class="color-field"><input id="color" type="color" value="#111827" /><code id="colorValue">#111827</code></span>
      </label>
      <label class="setting setting-wide">
        <span>数学字体</span>
        <select id="mathFont">
          <option value="tex">TeX 数学字体</option>
          <option value="times-new-roman">Times New Roman（MathType 风格）</option>
        </select>
      </label>
      <label class="setting setting-wide">
        <span>字体样式</span>
        <select id="mathStyle">
          <option value="default">数学默认（变量斜体）</option>
          <option value="upright">正体</option>
          <option value="italic">斜体</option>
          <option value="bold">粗体</option>
          <option value="bold-italic">粗斜体</option>
        </select>
      </label>
      <label class="setting">
        <span>排版</span>
        <select id="displayMode">
          <option value="display">独立公式</option>
          <option value="inline">行内公式</option>
        </select>
      </label>
      <label class="setting">
        <span>背景</span>
        <select id="background">
          <option value="transparent">透明</option>
          <option value="white">白色</option>
        </select>
      </label>
      <label class="setting setting-wide">
        <span>插入格式</span>
        <select id="insertFormat">
          <option value="auto">智能兼容（推荐）</option>
          <option value="omml" hidden>可编辑公式（仅 Word）</option>
          <option value="numbered" hidden>MathType 风格（独立公式自动编号）</option>
          <option value="svg">强制 SVG（实验性）</option>
          <option value="png">高清透明 PNG</option>
        </select>
        <small id="insertFormatHint" class="setting-hint" hidden></small>
      </label>
    </section>

    <button id="insertButton" class="primary-button" type="button">
      <span>插入为矢量公式</span><kbd>⌃⇧I</kbd>
    </button>
    <div class="secondary-actions">
      <button id="copyButton" type="button">复制 SVG</button>
      <button id="svgButton" type="button">下载 SVG</button>
      <button id="pngButton" type="button">下载高清 PNG</button>
    </div>
    <div id="status" class="message status" role="status" hidden></div>

    <footer>SVG 矢量优先 · 不支持时自动回退高清 PNG</footer>
  </main>
`;

const elements = {
  formula: document.querySelector<HTMLTextAreaElement>("#formula")!,
  fontSize: document.querySelector<HTMLInputElement>("#fontSize")!,
  sizeHint: document.querySelector<HTMLElement>("#sizeHint")!,
  color: document.querySelector<HTMLInputElement>("#color")!,
  colorValue: document.querySelector<HTMLElement>("#colorValue")!,
  mathFont: document.querySelector<HTMLSelectElement>("#mathFont")!,
  mathStyle: document.querySelector<HTMLSelectElement>("#mathStyle")!,
  displayMode: document.querySelector<HTMLSelectElement>("#displayMode")!,
  background: document.querySelector<HTMLSelectElement>("#background")!,
  insertFormat: document.querySelector<HTMLSelectElement>("#insertFormat")!,
  insertFormatHint: document.querySelector<HTMLElement>("#insertFormatHint")!,
  preview: document.querySelector<HTMLElement>("#preview")!,
  dimensions: document.querySelector<HTMLElement>("#dimensions")!,
  error: document.querySelector<HTMLElement>("#error")!,
  status: document.querySelector<HTMLElement>("#status")!,
  insert: document.querySelector<HTMLButtonElement>("#insertButton")!,
  editSelected: document.querySelector<HTMLButtonElement>("#editSelectedButton")!,
  hostBadge: document.querySelector<HTMLElement>("#hostBadge")!,
};

let rendered: RenderedFormula | null = null;
let renderTimer = 0;
let editingFormulaTarget: FormulaReplacementTarget | null = null;

function options(): RenderOptions {
  return {
    display: elements.displayMode.value === "display",
    color: elements.color.value,
    fontSizePt: Number(elements.fontSize.value) || 12,
    background: elements.background.value as RenderOptions["background"],
    mathFont: elements.mathFont.value as RenderOptions["mathFont"],
    mathStyle: elements.mathStyle.value as RenderOptions["mathStyle"],
  };
}

function showStatus(message: string, kind: "ok" | "warn" = "ok") {
  elements.status.textContent = message;
  elements.status.className = `message status ${kind}`;
  elements.status.hidden = false;
  window.setTimeout(() => (elements.status.hidden = true), 3500);
}

function formatPoint(value: number): string {
  return Number.isInteger(value) ? String(value) : value.toFixed(1).replace(/\.0$/, "");
}

function updateSizeHint() {
  const full = Number(elements.fontSize.value) || 12;
  elements.sizeHint.textContent =
    `脚标 ${formatPoint(full * MATHTYPE_SIZE_RATIOS.subscript)} pt · ` +
    `二级 ${formatPoint(full * MATHTYPE_SIZE_RATIOS.subSubscript)} pt`;
}

function updateInsertButtonLabel() {
  if (editingFormulaTarget) {
    elements.insert.querySelector("span")!.textContent = "更新选中公式";
    return;
  }
  const format = elements.insertFormat.value;
  const inlineNumbered = format === "numbered" && elements.displayMode.value === "inline";
  const label =
    format === "png"
      ? "插入高清公式"
      : format === "omml"
        ? "插入可编辑公式"
        : inlineNumbered
          ? "插入行内公式（无编号）"
          : format === "numbered"
          ? "插入排版公式并编号"
          : "插入为矢量公式";
  elements.insert.querySelector("span")!.textContent = label;

  if (format === "numbered") {
    elements.insertFormatHint.textContent = inlineNumbered
      ? "行内模式：在光标处紧凑插入，不添加前后空格或编号。"
      : "独立模式：公式居中，并在右侧使用 Word 域自动编号。";
    elements.insertFormatHint.hidden = false;
  } else {
    elements.insertFormatHint.hidden = true;
  }
}

function updatePreview() {
  if (!elements.formula.value.trim()) {
    rendered = null;
    elements.preview.innerHTML = '<span class="empty-preview">输入公式后在此预览…</span>';
    elements.dimensions.textContent = "";
    elements.error.hidden = true;
    elements.insert.disabled = true;
    localStorage.setItem("formula-canvas-source", "");
    return;
  }

  try {
    const renderOptions = options();
    rendered = renderFormula(elements.formula.value, renderOptions);
    elements.preview.innerHTML = rendered.svg;
    elements.dimensions.textContent =
      `Full ${formatPoint(renderOptions.fontSizePt)} pt · ` +
      `${rendered.widthPt.toFixed(1)} × ${rendered.heightPt.toFixed(1)} pt`;
    elements.error.hidden = true;
    elements.insert.disabled = false;
    localStorage.setItem("formula-canvas-source", elements.formula.value);
  } catch (error) {
    rendered = null;
    elements.preview.innerHTML = '<span class="empty-preview">等待有效公式…</span>';
    elements.dimensions.textContent = "";
    elements.error.textContent = error instanceof Error ? error.message : "公式解析失败";
    elements.error.hidden = false;
    elements.insert.disabled = true;
  }
}

function schedulePreview() {
  window.clearTimeout(renderTimer);
  renderTimer = window.setTimeout(updatePreview, 120);
}

function download(content: Blob, filename: string) {
  const url = URL.createObjectURL(content);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = filename;
  anchor.click();
  URL.revokeObjectURL(url);
}

async function handleInsert() {
  if (!rendered) return;
  elements.insert.disabled = true;
  elements.insert.querySelector("span")!.textContent = "正在插入…";
  try {
    const wasEditing = Boolean(editingFormulaTarget);
    const renderOptions = options();
    const inlineWithoutNumber =
      elements.insertFormat.value === "numbered" && elements.displayMode.value === "inline";
    const result = await insertFormula(
      rendered,
      elements.insertFormat.value as InsertPreference,
      { source: elements.formula.value, ...renderOptions },
      editingFormulaTarget,
    );
    editingFormulaTarget = null;
    if (wasEditing) {
      showStatus(
        result.sourceLinked
          ? "已更新选中的 Formula Canvas 公式"
          : "公式已更新，但当前 Word 版本无法保留源码关联",
        result.sourceLinked ? "ok" : "warn",
      );
    } else if (isWord() && !result.sourceLinked) {
      showStatus("公式已插入，但当前 Word 版本无法保留源码关联", "warn");
    } else if (inlineWithoutNumber) {
      showStatus(
        result.fallback
          ? "已插入行内公式（无编号；SVG 不可用，已使用高清 PNG）"
          : "已插入行内矢量公式（无编号、无额外空格）",
      );
    } else if (result.format === "OMML") {
      showStatus("已插入可编辑公式（Word 原生）");
    } else if (result.format === "NUMBERED") {
      showStatus("已插入 TeX 矢量公式和 Word 自动编号");
    } else if (result.format === "SVG") {
      showStatus(result.fallback ? "可编辑公式转换失败，已改用 SVG" : "已插入 SVG 矢量公式");
    } else if (result.fallback) {
      showStatus("插入失败，已自动改用高清 PNG");
    } else {
      showStatus("已用 PowerPoint 兼容模式插入高清透明 PNG");
    }
  } catch (error) {
    showStatus(error instanceof Error ? error.message : "插入失败", "warn");
  } finally {
    elements.insert.disabled = false;
    updateInsertButtonLabel();
  }
}

async function handleEditSelected() {
  elements.editSelected.disabled = true;
  elements.editSelected.textContent = "正在读取…";
  try {
    const formula = await recoverSelectedFormula();
    editingFormulaTarget = { tag: formula.tag, controlId: formula.controlId };
    elements.formula.value = formula.source;
    elements.displayMode.value = formula.display ? "display" : "inline";
    elements.fontSize.value = String(formula.fontSizePt);
    elements.color.value = formula.color;
    elements.colorValue.textContent = formula.color.toUpperCase();
    elements.background.value = formula.background;
    elements.mathFont.value = formula.mathFont;
    elements.mathStyle.value = formula.mathStyle;
    elements.insertFormat.value = formula.preference;
    localStorage.setItem("formula-canvas-full-size", elements.fontSize.value);
    localStorage.setItem("formula-canvas-math-font", elements.mathFont.value);
    localStorage.setItem("formula-canvas-math-style", elements.mathStyle.value);
    updateSizeHint();
    updatePreview();
    updateInsertButtonLabel();
    elements.formula.focus();
    showStatus("已恢复原 LaTeX 和排版设置；插入时将替换 Word 中的原公式");
  } catch (error) {
    showStatus(error instanceof Error ? error.message : "无法读取选中的公式", "warn");
  } finally {
    elements.editSelected.disabled = false;
    elements.editSelected.textContent = "编辑选中公式";
  }
}

function cancelFormulaEditing() {
  if (!editingFormulaTarget) return;
  editingFormulaTarget = null;
  updateInsertButtonLabel();
}

elements.formula.value = localStorage.getItem("formula-canvas-source") ?? SAMPLE;
elements.fontSize.value = localStorage.getItem("formula-canvas-full-size") ?? "12";
elements.mathFont.value = localStorage.getItem("formula-canvas-math-font") ?? "tex";
elements.mathStyle.value = localStorage.getItem("formula-canvas-math-style") ?? "default";
elements.formula.addEventListener("input", schedulePreview);
elements.background.addEventListener("input", schedulePreview);
elements.fontSize.addEventListener("input", () => {
  localStorage.setItem("formula-canvas-full-size", elements.fontSize.value);
  updateSizeHint();
  schedulePreview();
});
elements.displayMode.addEventListener("input", () => {
  schedulePreview();
  updateInsertButtonLabel();
});
elements.mathFont.addEventListener("input", () => {
  localStorage.setItem("formula-canvas-math-font", elements.mathFont.value);
  schedulePreview();
});
elements.mathStyle.addEventListener("input", () => {
  localStorage.setItem("formula-canvas-math-style", elements.mathStyle.value);
  schedulePreview();
});
elements.color.addEventListener("input", () => {
  elements.colorValue.textContent = elements.color.value.toUpperCase();
  schedulePreview();
});
elements.insertFormat.addEventListener("change", updateInsertButtonLabel);
document.querySelector("#sampleButton")!.addEventListener("click", () => {
  cancelFormulaEditing();
  elements.formula.value = SAMPLE;
  updatePreview();
  elements.formula.focus();
});
function clearFormula() {
  cancelFormulaEditing();
  elements.formula.value = "";
  updatePreview();
  elements.formula.focus();
}

document.querySelector("#clearButton")!.addEventListener("click", clearFormula);
elements.insert.addEventListener("click", handleInsert);
elements.editSelected.addEventListener("click", handleEditSelected);
document.addEventListener("keydown", (event) => {
  if (event.ctrlKey && event.shiftKey && !event.metaKey && !event.altKey) {
    const key = event.key.toLowerCase();
    if (key !== "i" && key !== "k") return;
    event.preventDefault();
    if (event.repeat) return;
    if (key === "i") void handleInsert();
    else clearFormula();
  }
});
document.querySelector("#copyButton")!.addEventListener("click", async () => {
  if (!rendered) return;
  await navigator.clipboard.writeText(rendered.svg);
  showStatus("SVG 源码已复制");
});
document.querySelector("#svgButton")!.addEventListener("click", () => {
  if (rendered) download(new Blob([rendered.svg], { type: "image/svg+xml" }), "formula.svg");
});
document.querySelector("#pngButton")!.addEventListener("click", async () => {
  if (!rendered) return;
  const base64 = await svgToPngBase64(rendered, 4);
  const response = await fetch(`data:image/png;base64,${base64}`);
  download(await response.blob(), "formula@4x.png");
});

function initializeHost() {
  const host = officeHostName();
  elements.hostBadge.textContent = host;
  elements.hostBadge.classList.toggle("connected", isInsideOffice());
  elements.insert.title = isInsideOffice()
    ? "插入到当前光标或幻灯片（Control + Shift + I）"
    : "请在 Office 中打开插件";

  // 可编辑公式（OMML）仅 Word 支持，其它宿主隐藏该选项。
  const ommlOption = elements.insertFormat.querySelector<HTMLOptionElement>('option[value="omml"]');
  const numberedOption =
    elements.insertFormat.querySelector<HTMLOptionElement>('option[value="numbered"]');
  if (ommlOption) {
    ommlOption.hidden = !isWord();
    if (!isWord() && ["omml", "numbered"].includes(elements.insertFormat.value)) {
      elements.insertFormat.value = "auto";
    }
  }
  if (numberedOption) numberedOption.hidden = !isWord();
  const canRoundTrip = supportsFormulaRoundTrip();
  elements.editSelected.hidden = !canRoundTrip;
  elements.editSelected.title = canRoundTrip
    ? "恢复当前选中的 Formula Canvas 公式"
    : "当前 Word 版本不支持恢复公式源码";

  updateInsertButtonLabel();
}

function isEmbeddedSiteFrame() {
  if (window.self === window.top) return false;
  try {
    const referrer = new URL(document.referrer);
    return referrer.origin === window.location.origin && referrer.pathname.startsWith("/site/");
  } catch {
    return false;
  }
}

function loadOfficeJs() {
  return new Promise<void>((resolve, reject) => {
    if (typeof Office !== "undefined") {
      resolve();
      return;
    }
    const script = document.createElement("script");
    script.src = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";
    script.onload = () => resolve();
    script.onerror = () => reject(new Error("Office.js 加载失败"));
    document.head.append(script);
  });
}

async function initializeOfficeIntegration() {
  if (isEmbeddedSiteFrame()) {
    initializeHost();
    return;
  }

  try {
    await loadOfficeJs();
    if (typeof Office !== "undefined") Office.onReady(initializeHost);
    else initializeHost();
  } catch {
    initializeHost();
  }
}

updateSizeHint();
updatePreview();
void initializeOfficeIntegration();
window.setTimeout(initializeHost, 1500);
