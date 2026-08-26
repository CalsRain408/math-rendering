import { mathjax } from "mathjax-full/js/mathjax.js";
import { TeX } from "mathjax-full/js/input/tex.js";
import { SVG } from "mathjax-full/js/output/svg.js";
import { liteAdaptor } from "mathjax-full/js/adaptors/liteAdaptor.js";
import { RegisterHTMLHandler } from "mathjax-full/js/handlers/html.js";
import { AllPackages } from "mathjax-full/js/input/tex/AllPackages.js";
import { SerializedMmlVisitor } from "mathjax-full/js/core/MmlTree/SerializedMmlVisitor.js";
import { CommonWrapper } from "mathjax-full/js/output/common/Wrapper.js";

export const MATHTYPE_SIZE_RATIOS = {
  subscript: 0.58,
  subSubscript: 0.42,
} as const;

/**
 * MathJax's TeX defaults use 70.7% scripts and 50% nested scripts. MathType's
 * factory size definitions use 58% and 42% of Full size. Override only the
 * SVG wrapper scale calculation so layout metrics and serialized transforms
 * are computed with the MathType ratios from the start.
 */
function configureMathTypeSizing(): void {
  const prototype = CommonWrapper.prototype as unknown as {
    formulaCanvasMathTypeSizing?: boolean;
    getScale(this: any): void;
  };
  if (prototype.formulaCanvasMathTypeSizing) return;

  prototype.getScale = function (this: any) {
    let scale = 1;
    const parentScale = this.parent ? this.parent.bbox.scale : 1;
    const attributes = this.node.attributes;
    const scriptLevel = Math.max(0, Math.min(Number(attributes.get("scriptlevel")), 2));
    let fontSize = attributes.get("fontsize");
    let mathSize =
      this.node.isToken || this.node.isKind("mstyle")
        ? attributes.get("mathsize")
        : attributes.getInherited("mathsize");

    if (scriptLevel === 1) scale = MATHTYPE_SIZE_RATIOS.subscript;
    else if (scriptLevel >= 2) scale = MATHTYPE_SIZE_RATIOS.subSubscript;

    if (this.removedStyles?.fontSize && !fontSize) fontSize = this.removedStyles.fontSize;
    if (fontSize && !attributes.getExplicit("mathsize")) mathSize = fontSize;
    if (mathSize !== "1") scale *= this.length2em(mathSize, 1, 1);

    this.bbox.scale = scale;
    this.bbox.rscale = scale / parentScale;
  };
  prototype.formulaCanvasMathTypeSizing = true;
}

configureMathTypeSizing();

export type RenderOptions = {
  display: boolean;
  color: string;
  fontSizePt: number;
  background: "transparent" | "white";
  mathFont: MathFont;
  mathStyle: MathStyle;
};

export type MathFont = "tex" | "times-new-roman";
export type MathStyle = "default" | "upright" | "italic" | "bold" | "bold-italic";

export type RenderedFormula = {
  svg: string;
  widthPx: number;
  heightPx: number;
  widthPt: number;
  heightPt: number;
  /** Distance from the bottom of the image to MathJax's math baseline. */
  baselineDepthPt: number;
};

const adaptor = liteAdaptor();
RegisterHTMLHandler(adaptor);

const tex = new TeX({
  packages: AllPackages,
  formatError: (_jax: unknown, error: Error) => {
    throw error;
  },
});
const svgOutput = new SVG({ fontCache: "none" });
const document = mathjax.document("", { InputJax: tex, OutputJax: svgOutput });
const mmlVisitor = new SerializedMmlVisitor();

function parseEx(value: string | null, fallback: number): number {
  if (!value) return fallback;
  const parsed = Number.parseFloat(value);
  return Number.isFinite(parsed) ? parsed : fallback;
}

function verticalAlignEx(style: string | null): number {
  if (!style) return 0;
  const match = /(?:^|;)\s*vertical-align\s*:\s*(-?(?:\d+(?:\.\d*)?|\.\d+))ex\b/i.exec(
    style,
  );
  if (!match) return 0;
  const parsed = Number.parseFloat(match[1]);
  return Number.isFinite(parsed) ? parsed : 0;
}

function normalizeMathInput(source: string): string {
  const normalized = source.trim();
  const delimiterPairs: Array<[string, string]> = [
    ["$$", "$$"],
    ["\\[", "\\]"],
    ["\\(", "\\)"],
    ["$", "$"],
  ];
  for (const [opening, closing] of delimiterPairs) {
    if (
      normalized.length >= opening.length + closing.length &&
      normalized.startsWith(opening) &&
      normalized.endsWith(closing)
    ) {
      return normalized.slice(opening.length, -closing.length).trim();
    }
  }
  return normalized;
}

function expandUseElements(svg: SVGSVGElement): void {
  const definitions = new Map<string, Element>();
  svg.querySelectorAll("[id]").forEach((element) => {
    if (element.id) definitions.set(element.id, element);
  });

  svg.querySelectorAll("use").forEach((use) => {
    const href =
      use.getAttribute("href") ||
      use.getAttribute("xlink:href") ||
      use.getAttributeNS("http://www.w3.org/1999/xlink", "href");
    if (!href?.startsWith("#")) return;

    const definition = definitions.get(href.slice(1));
    if (!definition) return;

    const replacement = definition.cloneNode(true) as Element;
    replacement.removeAttribute("id");

    const transforms: string[] = [];
    const useTransform = use.getAttribute("transform");
    const x = use.getAttribute("x");
    const y = use.getAttribute("y");
    const originalTransform = replacement.getAttribute("transform");
    if (useTransform) transforms.push(useTransform);
    if (x || y) transforms.push(`translate(${x || "0"} ${y || "0"})`);
    if (originalTransform) transforms.push(originalTransform);
    if (transforms.length) replacement.setAttribute("transform", transforms.join(" "));

    for (const attribute of Array.from(use.attributes)) {
      if (["href", "xlink:href", "x", "y", "transform"].includes(attribute.name)) continue;
      if (!attribute.name.startsWith("data-") && !attribute.name.startsWith("aria-")) {
        replacement.setAttribute(attribute.name, attribute.value);
      }
    }
    use.replaceWith(replacement);
  });

  svg.querySelectorAll("defs").forEach((defs) => defs.remove());
}

type TimesGlyph = {
  character: string;
  italic: boolean;
  bold: boolean;
};

function alphabetGlyph(
  codePoint: number,
  start: number,
  end: number,
  firstCharacter: string,
  italic: boolean,
  bold: boolean,
): TimesGlyph | null {
  if (codePoint < start || codePoint > end) return null;
  return {
    character: String.fromCodePoint(firstCharacter.codePointAt(0)! + codePoint - start),
    italic,
    bold,
  };
}

/** Map MathJax's Unicode math-alphabet glyph IDs back to Times characters. */
export function timesNewRomanGlyph(codePoint: number): TimesGlyph | null {
  if (codePoint >= 0x30 && codePoint <= 0x39) {
    return { character: String.fromCodePoint(codePoint), italic: false, bold: false };
  }
  if (codePoint >= 0x1d7ce && codePoint <= 0x1d7d7) {
    return {
      character: String.fromCodePoint(0x30 + codePoint - 0x1d7ce),
      italic: false,
      bold: true,
    };
  }
  if ((codePoint >= 0x41 && codePoint <= 0x5a) || (codePoint >= 0x61 && codePoint <= 0x7a)) {
    return { character: String.fromCodePoint(codePoint), italic: false, bold: false };
  }

  const ranges: Array<[number, number, string, boolean, boolean]> = [
    [0x1d400, 0x1d419, "A", false, true],
    [0x1d41a, 0x1d433, "a", false, true],
    [0x1d434, 0x1d44d, "A", true, false],
    [0x1d44e, 0x1d467, "a", true, false],
    [0x1d468, 0x1d481, "A", true, true],
    [0x1d482, 0x1d49b, "a", true, true],
  ];
  for (const range of ranges) {
    const glyph = alphabetGlyph(codePoint, ...range);
    if (glyph) return glyph;
  }

  // Mathematical italic small h is encoded outside the contiguous alphabet.
  if (codePoint === 0x210e) return { character: "h", italic: true, bold: false };
  return null;
}

/**
 * MathType traditionally combines Times letters/numbers with dedicated math
 * glyphs for operators and stretchy constructions. Replacing only token glyphs
 * keeps MathJax's fraction, root, integral, spacing, and script layout intact.
 */
function applyTimesNewRoman(svg: SVGSVGElement, mathStyle: MathStyle): void {
  const namespace = "http://www.w3.org/2000/svg";
  let tokenId = 0;
  const tokenIds = new Map<SVGGElement, string>();
  svg.querySelectorAll<SVGPathElement>("path[data-c]").forEach((path) => {
    const token = path.closest<SVGGElement>("g[data-mml-node]");
    if (!token) return;
    const tokenType = token?.getAttribute("data-mml-node");
    if (!tokenType || !["mi", "mn", "mtext"].includes(tokenType)) return;

    const code = path.getAttribute("data-c");
    if (!code) return;
    const glyph = timesNewRomanGlyph(Number.parseInt(code, 16));
    if (!glyph) return;
    if (tokenType === "mn" && !/\d/.test(glyph.character)) return;

    const text = globalThis.document.createElementNS(namespace, "text");
    text.textContent = glyph.character;
    text.setAttribute("x", "0");
    text.setAttribute("y", "0");
    const styledGlyph =
      mathStyle === "default"
        ? glyph
        : {
            character: glyph.character,
            italic: mathStyle === "italic" || mathStyle === "bold-italic",
            bold: mathStyle === "bold" || mathStyle === "bold-italic",
          };

    text.setAttribute("font-family", "Times New Roman, Times, serif");
    text.setAttribute("font-size", "1000");
    let id = tokenIds.get(token);
    if (!id) {
      id = String(++tokenId);
      tokenIds.set(token, id);
    }
    text.setAttribute("data-times-token", id);
    text.setAttribute("data-times-token-type", tokenType);
    if (styledGlyph.italic) text.setAttribute("font-style", "italic");
    if (styledGlyph.bold) text.setAttribute("font-weight", "bold");

    // MathJax paths use an upward Y axis. SVG text uses the opposite axis, so
    // flip the text locally while retaining any per-glyph translation.
    const transform = path.getAttribute("transform");
    text.setAttribute("transform", `${transform ? `${transform} ` : ""}scale(1,-1)`);
    path.replaceWith(text);
  });
}

type TimesTextCandidate = {
  container: Element;
  text: SVGTextElement;
  tokenType: string;
  x: number;
  y: number;
  remainingTransform: string;
  signature: string;
};

function leadingTranslation(
  element: Element,
): { x: number; y: number; remainingTransform: string } | null {
  const transform = element.getAttribute("transform");
  if (!transform) return { x: 0, y: 0, remainingTransform: "" };
  const match =
    /^\s*translate\(\s*(-?(?:\d+(?:\.\d*)?|\.\d+)(?:e[-+]?\d+)?)(?:[\s,]+(-?(?:\d+(?:\.\d*)?|\.\d+)(?:e[-+]?\d+)?))?\s*\)/i.exec(
      transform,
    );
  if (!match) return { x: 0, y: 0, remainingTransform: transform.trim() };
  return {
    x: Number.parseFloat(match[1]),
    y: match[2] ? Number.parseFloat(match[2]) : 0,
    remainingTransform: transform.slice(match[0].length).trim(),
  };
}

function timesTextCandidate(element: Element): TimesTextCandidate | null {
  let text: SVGTextElement;
  if (element.localName === "text" && element.hasAttribute("data-times-token")) {
    text = element as SVGTextElement;
  } else if (element.localName === "g") {
    const texts = element.querySelectorAll<SVGTextElement>("text[data-times-token]");
    if (texts.length !== 1 || element.querySelector("path, use, rect")) return null;
    text = texts[0];
  } else {
    return null;
  }

  const translation = leadingTranslation(element);
  const tokenType = text.getAttribute("data-times-token-type");
  if (!translation || !tokenType || !["mi", "mn", "mtext"].includes(tokenType)) return null;

  const signature = [
    tokenType,
    text.getAttribute("font-family") || "",
    text.getAttribute("font-size") || "",
    text.getAttribute("font-style") || "",
    text.getAttribute("font-weight") || "",
    element === text ? translation.remainingTransform : text.getAttribute("transform") || "",
  ].join("|");
  return { container: element, text, tokenType, ...translation, signature };
}

function horizontalScaleInParent(text: SVGTextElement, parent: SVGGElement): number {
  try {
    const textMatrix = text.getCTM();
    const parentMatrix = parent.getCTM();
    if (!textMatrix || !parentMatrix) return 1;
    const relative = parentMatrix.inverse().multiply(textMatrix);
    const scale = Math.hypot(relative.a, relative.b);
    return Number.isFinite(scale) && scale > 0 ? scale : 1;
  } catch {
    return 1;
  }
}

function shiftFollowingSiblings(start: Element | null, delta: number): void {
  for (let sibling = start; sibling; sibling = sibling.nextElementSibling) {
    const transform = sibling.getAttribute("transform");
    sibling.setAttribute(
      "transform",
      `translate(${delta.toFixed(4)} 0)${transform ? ` ${transform}` : ""}`,
    );
  }
}

function screenScaleX(element: SVGGraphicsElement): number {
  try {
    const matrix = element.getScreenCTM();
    if (!matrix) return 0;
    const scale = Math.hypot(matrix.a, matrix.b);
    return Number.isFinite(scale) && scale > 0 ? scale : 0;
  } catch {
    return 0;
  }
}

/** Move the first following item in an ancestor row by a screen-space delta. */
function propagateFollowingShift(parent: SVGGElement, deltaPx: number): void {
  let current: Element = parent;
  while (current.parentElement?.localName === "g") {
    const ancestor = current.parentElement as unknown as SVGGElement;
    const next = current.nextElementSibling;
    if (next) {
      const scale = screenScaleX(ancestor);
      if (scale > 0) shiftFollowingSiblings(next, deltaPx / scale);
      return;
    }
    current = ancestor;
  }
}

function suppressLatinLigatures(value: string): string {
  // ZWNJ is zero-width but reliably prevents fi/fl/ffi/ffl shaping even in
  // SVG consumers that ignore the CSS font-variant-ligatures declaration.
  return value.replace(/f(?=[fil])/g, "f\u200c");
}

/**
 * MathJax positions every TeX path glyph independently. After swapping those
 * paths for Times text, the old TeX advances can leave gaps (notably M → a) or
 * overlaps. Join adjacent Times tokens into native SVG text runs so the Times
 * font performs its own kerning, then move the remaining row by the exact
 * change in advance. MathJax still controls operators, scripts and fractions.
 */
function compactTimesTextRuns(svg: SVGSVGElement): number {
  const viewBox = svg
    .getAttribute("viewBox")
    ?.trim()
    .split(/[\s,]+/)
    .map(Number);
  if (!viewBox || viewBox.length !== 4 || viewBox.some((value) => !Number.isFinite(value))) {
    return 0;
  }

  const measurementHost = globalThis.document.createElement("div");
  measurementHost.style.cssText =
    "position:absolute;left:-100000px;top:-100000px;visibility:hidden;" +
    "pointer-events:none;overflow:visible;contain:none;";
  const measuredSvg = globalThis.document.importNode(svg, true) as SVGSVGElement;
  measuredSvg.style.cssText = "display:block;max-width:none;overflow:visible;";
  measurementHost.appendChild(measuredSvg);
  globalThis.document.body.appendChild(measurementHost);

  try {
    const viewport = measuredSvg.getBoundingClientRect();
    if (viewport.width <= 0) return 0;
    const content = Array.from(measuredSvg.children).filter((element) => element.localName !== "rect");
    const originalRight = Math.max(...content.map((element) => element.getBoundingClientRect().right));

    const parents = Array.from(measuredSvg.querySelectorAll<SVGGElement>("g")).reverse();
    for (const parent of parents) {
      const children = Array.from(parent.children);
      const runs: TimesTextCandidate[][] = [];
      let current: TimesTextCandidate[] = [];

      const finishRun = () => {
        if (current.length >= 2) runs.push(current);
        current = [];
      };

      for (const child of children) {
        const candidate = timesTextCandidate(child);
        const previous = current[current.length - 1];
        if (
          candidate &&
          (!previous ||
            (candidate.signature === previous.signature &&
              Math.abs(candidate.y - previous.y) < 0.001 &&
              candidate.x > previous.x))
        ) {
          current.push(candidate);
        } else {
          finishRun();
          if (candidate) current.push(candidate);
        }
      }
      finishRun();

      // Work right-to-left so moving later siblings cannot invalidate the
      // original coordinates used to compact an earlier run in the same row.
      for (const run of runs.reverse()) {
        const first = run[0];
        const last = run[run.length - 1];
        const next = last.container.nextElementSibling;
        const nextTranslation = next ? leadingTranslation(next) : null;
        const originalInkRightPx = Math.max(
          ...run.map((item) => item.container.getBoundingClientRect().right),
        );

        first.text.textContent = suppressLatinLigatures(
          run.map((item) => item.text.textContent || "").join(""),
        );
        first.text.setAttribute("font-kerning", "normal");
        first.text.setAttribute("letter-spacing", "0");
        first.text.setAttribute("font-variant-ligatures", "none");
        first.text.setAttribute("font-feature-settings", "'liga' 0, 'clig' 0, 'dlig' 0");
        first.text.setAttribute(
          "style",
          "font-variant-ligatures:none;font-feature-settings:'liga' 0,'clig' 0,'dlig' 0",
        );
        for (const item of run.slice(1)) item.container.remove();

        if (!next || !nextTranslation) {
          const compactedInkRightPx = first.text.getBoundingClientRect().right;
          const deltaPx = compactedInkRightPx - originalInkRightPx;
          if (Number.isFinite(deltaPx) && Math.abs(deltaPx) >= 0.01) {
            propagateFollowingShift(parent, deltaPx);
          }
          continue;
        }
        const nativeAdvance =
          first.text.getComputedTextLength() * horizontalScaleInParent(first.text, parent);
        const originalAdvance = nextTranslation.x - first.x;
        const delta = nativeAdvance - originalAdvance;
        if (Number.isFinite(delta) && Math.abs(delta) >= 0.01) {
          shiftFollowingSiblings(next, delta);
        }
      }
    }

    const compactedRight = Math.max(...content.map((element) => element.getBoundingClientRect().right));
    // Keep a half-pixel guard for Word/Chromium font-metric rounding while
    // removing the obsolete right-side space created by the TeX advances.
    const rightDeltaPx = Math.min(0, compactedRight - originalRight + 0.5);
    const [, , width] = viewBox;
    if (rightDeltaPx < -0.01) {
      const adjustedWidth = Math.max(1, width + rightDeltaPx * (width / viewport.width));
      svg.setAttribute("viewBox", `${viewBox[0]} ${viewBox[1]} ${adjustedWidth} ${viewBox[3]}`);
    }

    svg.replaceChildren(
      ...Array.from(measuredSvg.childNodes).map((node) => svg.ownerDocument.importNode(node, true)),
    );
    return rightDeltaPx;
  } catch (error) {
    console.warn("Times New Roman text-run compaction failed", error);
    return 0;
  } finally {
    measurementHost.remove();
  }
}

type ViewportExpansion = {
  leftPx: number;
  rightPx: number;
  topPx: number;
  bottomPx: number;
};

/**
 * MathJax's viewBox is calculated for its TeX path glyphs. Times New Roman
 * italic letters (especially f) can have negative side bearings or taller ink,
 * so replacing a path with live SVG text may otherwise clip the overhang.
 */
function expandViewportForTimesText(svg: SVGSVGElement): ViewportExpansion {
  const empty: ViewportExpansion = { leftPx: 0, rightPx: 0, topPx: 0, bottomPx: 0 };
  const viewBox = svg
    .getAttribute("viewBox")
    ?.trim()
    .split(/[\s,]+/)
    .map(Number);
  if (!viewBox || viewBox.length !== 4 || viewBox.some((value) => !Number.isFinite(value))) {
    return empty;
  }

  const measurementHost = globalThis.document.createElement("div");
  measurementHost.style.cssText =
    "position:absolute;left:-100000px;top:-100000px;visibility:hidden;" +
    "pointer-events:none;overflow:visible;contain:none;";
  const measuredSvg = globalThis.document.importNode(svg, true) as SVGSVGElement;
  measuredSvg.style.cssText = "display:block;max-width:none;overflow:visible;";
  measurementHost.appendChild(measuredSvg);
  globalThis.document.body.appendChild(measurementHost);

  try {
    const viewport = measuredSvg.getBoundingClientRect();
    const contentElements = Array.from(measuredSvg.children).filter(
      (element) => element.localName !== "rect",
    );
    if (!contentElements.length || viewport.width <= 0 || viewport.height <= 0) return empty;

    let inkLeft = viewport.right;
    let inkRight = viewport.left;
    let inkTop = viewport.bottom;
    let inkBottom = viewport.top;
    for (const element of contentElements) {
      // Measuring the complete rendered group catches not only italic f, but
      // also Times capitals and superscripts whose ink can exceed MathJax's
      // original TeX path viewport in several directions.
      const bounds = element.getBoundingClientRect();
      inkLeft = Math.min(inkLeft, bounds.left);
      inkRight = Math.max(inkRight, bounds.right);
      inkTop = Math.min(inkTop, bounds.top);
      inkBottom = Math.max(inkBottom, bounds.bottom);
    }

    // A small guard absorbs sub-pixel rounding and minor Times font metric
    // differences between the browser renderer and Word's SVG renderer.
    const guardPx = 0.5;
    const expansion: ViewportExpansion = {
      leftPx: inkLeft < viewport.left ? viewport.left - inkLeft + guardPx : 0,
      rightPx: inkRight > viewport.right ? inkRight - viewport.right + guardPx : 0,
      topPx: inkTop < viewport.top ? viewport.top - inkTop + guardPx : 0,
      bottomPx: inkBottom > viewport.bottom ? inkBottom - viewport.bottom + guardPx : 0,
    };
    if (!Object.values(expansion).some((value) => value > 0)) return empty;

    const [x, y, width, height] = viewBox;
    const unitsPerPxX = width / viewport.width;
    const unitsPerPxY = height / viewport.height;
    const left = expansion.leftPx * unitsPerPxX;
    const right = expansion.rightPx * unitsPerPxX;
    const top = expansion.topPx * unitsPerPxY;
    const bottom = expansion.bottomPx * unitsPerPxY;
    svg.setAttribute(
      "viewBox",
      `${x - left} ${y - top} ${width + left + right} ${height + top + bottom}`,
    );
    return expansion;
  } catch (error) {
    console.warn("Times New Roman glyph bounds measurement failed", error);
    return empty;
  } finally {
    measurementHost.remove();
  }
}

function styledFormulaSource(source: string, mathStyle: MathStyle): string {
  switch (mathStyle) {
    case "upright":
      return `\\mathrm{${source}}`;
    case "italic":
      return `\\mathit{${source}}`;
    case "bold":
      return `\\mathbf{${source}}`;
    case "bold-italic":
      return `\\boldsymbol{\\mathit{${source}}}`;
    default:
      return source;
  }
}

function makeOfficeSafe(svg: SVGSVGElement, color: string): string {
  expandUseElements(svg);

  svg.querySelectorAll("*").forEach((element) => {
    for (const attribute of Array.from(element.attributes)) {
      if (
        (attribute.name === "style" &&
          !(
            element.localName === "text" &&
            attribute.value.includes("font-variant-ligatures:none")
          )) ||
        attribute.name === "class" ||
        attribute.name.startsWith("data-") ||
        attribute.name.startsWith("aria-")
      ) {
        element.removeAttribute(attribute.name);
        continue;
      }
      if (attribute.value.includes("currentColor")) {
        element.setAttribute(attribute.name, attribute.value.replaceAll("currentColor", color));
      }
    }
  });

  svg.removeAttribute("style");
  svg.removeAttribute("class");
  svg.removeAttribute("xmlns:xlink");
  svg.setAttribute("xmlns", "http://www.w3.org/2000/svg");
  svg.setAttribute("version", "1.1");
  svg.setAttribute("fill", color);
  svg.setAttribute("color", color);
  svg.setAttribute("shape-rendering", "geometricPrecision");

  const serialized = new XMLSerializer().serializeToString(svg);
  const parsed = new DOMParser().parseFromString(serialized, "image/svg+xml");
  if (parsed.querySelector("parsererror") || parsed.documentElement.localName !== "svg") {
    throw new Error("生成的 SVG 未通过 Office 兼容性检查");
  }
  return new XMLSerializer().serializeToString(parsed.documentElement);
}

export function renderFormula(source: string, options: RenderOptions): RenderedFormula {
  const normalized = normalizeMathInput(source);
  if (!normalized) throw new Error("请输入公式");

  const node = document.convert(styledFormulaSource(normalized, options.mathStyle), {
    display: options.display,
  });
  const wrapper = new DOMParser().parseFromString(adaptor.outerHTML(node), "text/html");
  const svg = wrapper.querySelector("svg");
  if (!svg) throw new Error("未能生成 SVG");

  const widthEx = parseEx(svg.getAttribute("width"), 8);
  const heightEx = parseEx(svg.getAttribute("height"), 2);
  // MathJax expresses the outer dimensions in ex. Convert through the actual
  // font x-height to MathType's Full size in points, then to CSS px for preview.
  const xHeight = svgOutput.font.params.x_height;
  let widthPt = Math.max(0.5, widthEx * xHeight * options.fontSizePt);
  let heightPt = Math.max(0.5, heightEx * xHeight * options.fontSizePt);
  // MathJax places the visual baseline above the SVG bottom using a negative
  // CSS vertical-align value. Office inserts an image with its bottom on the
  // text baseline, so retain this depth for Word to lower inline formulas by
  // the exact amount required (not a fixed, formula-specific guess).
  let baselineDepthPt = Math.max(
    0,
    -verticalAlignEx(svg.getAttribute("style")) * xHeight * options.fontSizePt,
  );
  let widthPx = (widthPt * 96) / 72;
  let heightPx = (heightPt * 96) / 72;

  svg.setAttribute("width", widthPx.toFixed(2));
  svg.setAttribute("height", heightPx.toFixed(2));
  svg.setAttribute("preserveAspectRatio", "xMidYMid meet");
  svg.removeAttribute("focusable");
  svg.removeAttribute("role");

  if (options.background === "white") {
    const rect = wrapper.createElementNS("http://www.w3.org/2000/svg", "rect");
    rect.setAttribute("x", "0");
    rect.setAttribute("y", "0");
    rect.setAttribute("width", "100%");
    rect.setAttribute("height", "100%");
    rect.setAttribute("fill", "white");
    svg.insertBefore(rect, svg.firstChild);
  }

  if (options.mathFont === "times-new-roman") {
    applyTimesNewRoman(svg, options.mathStyle);
    const compactedWidthPx = compactTimesTextRuns(svg);
    widthPx += compactedWidthPx;
    widthPt = (widthPx * 72) / 96;
    svg.setAttribute("width", widthPx.toFixed(2));
    const expansion = expandViewportForTimesText(svg);
    widthPx += expansion.leftPx + expansion.rightPx;
    heightPx += expansion.topPx + expansion.bottomPx;
    baselineDepthPt += (expansion.bottomPx * 72) / 96;
    widthPt = (widthPx * 72) / 96;
    heightPt = (heightPx * 72) / 96;
    svg.setAttribute("width", widthPx.toFixed(2));
    svg.setAttribute("height", heightPx.toFixed(2));
  }

  return {
    svg: makeOfficeSafe(svg, options.color),
    widthPx,
    heightPx,
    widthPt,
    heightPt,
    baselineDepthPt,
  };
}

/**
 * 把 LaTeX 转成 MathML 字符串（供 OMML 可编辑公式使用）。
 * 复用与 SVG 渲染相同的 TeX 解析器，保证语法支持一致。
 */
export function renderMathML(source: string, display: boolean): string {
  const normalized = normalizeMathInput(source);
  if (!normalized) throw new Error("请输入公式");

  const MathItem = document.options.MathItem;
  const math = new MathItem(normalized, tex, display);
  math.compile(document);
  return mmlVisitor.visitTree(math.root);
}

export async function svgToPngBase64(
  rendered: RenderedFormula,
  scale = 4,
): Promise<string> {
  const blob = new Blob([rendered.svg], { type: "image/svg+xml;charset=utf-8" });
  const url = URL.createObjectURL(blob);
  try {
    const image = new Image();
    image.decoding = "async";
    image.src = url;
    await image.decode();

    const canvas = globalThis.document.createElement("canvas");
    canvas.width = Math.ceil(rendered.widthPx * scale);
    canvas.height = Math.ceil(rendered.heightPx * scale);
    const context = canvas.getContext("2d");
    if (!context) throw new Error("浏览器无法创建 Canvas");
    context.scale(scale, scale);
    context.drawImage(image, 0, 0, rendered.widthPx, rendered.heightPx);
    return canvas.toDataURL("image/png").split(",")[1];
  } finally {
    URL.revokeObjectURL(url);
  }
}
