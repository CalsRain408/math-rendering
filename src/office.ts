import {
  DEFAULT_MATHTYPE_SIZE_RATIOS,
  svgToPngBase64,
  renderMathML,
  type RenderOptions,
  type RenderedFormula,
} from "./renderer";
import { mml2omml } from "mathml2omml";

export type InsertResult = {
  format: "SVG" | "PNG" | "OMML" | "NUMBERED";
  fallback: boolean;
  sourceLinked: boolean;
};
export type InsertPreference = "auto" | "svg" | "png" | "omml" | "numbered";
export type FormulaEditData = RenderOptions & {
  source: string;
  preference: InsertPreference;
};
export type FormulaReplacementTarget = { tag: string; controlId: number };
export type RecoveredFormula = FormulaEditData & FormulaReplacementTarget;

const FORMULA_TAG_PREFIX = "FormulaCanvas:";
const FORMULA_METADATA_NAMESPACE = "urn:formula-canvas:formula:v1";

type CustomXmlPartLike = {
  id: string;
  getXmlAsync(callback: (result: Office.AsyncResult<string>) => void): void;
  deleteAsync(callback: (result: Office.AsyncResult<void>) => void): void;
};

function customXmlParts(): any | null {
  if (!isInsideOffice()) return null;
  return (Office.context.document as any).customXmlParts ?? null;
}

function base64ToUtf8(value: string): string {
  const binary = atob(value);
  const bytes = Uint8Array.from(binary, (character) => character.charCodeAt(0));
  return new TextDecoder().decode(bytes);
}

function customXmlResultError(result: Office.AsyncResult<unknown>, fallback: string): Error {
  return new Error(result.error?.message || fallback);
}

function addCustomXmlPart(xml: string): Promise<CustomXmlPartLike> {
  return new Promise((resolve, reject) => {
    const parts = customXmlParts();
    if (!parts) {
      reject(new Error("当前 Word 版本不支持公式源码元数据"));
      return;
    }
    parts.addAsync(xml, (result: Office.AsyncResult<CustomXmlPartLike>) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        resolve(result.value);
      } else {
        reject(customXmlResultError(result, "无法保存公式源码元数据"));
      }
    });
  });
}

function getCustomXmlPart(id: string): Promise<CustomXmlPartLike> {
  return new Promise((resolve, reject) => {
    const parts = customXmlParts();
    if (!parts) {
      reject(new Error("当前 Word 版本不支持读取公式源码元数据"));
      return;
    }
    parts.getByIdAsync(id, (result: Office.AsyncResult<CustomXmlPartLike>) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && result.value) {
        resolve(result.value);
      } else {
        reject(customXmlResultError(result, "没有找到该公式的源码元数据"));
      }
    });
  });
}

function readCustomXml(part: CustomXmlPartLike): Promise<string> {
  return new Promise((resolve, reject) => {
    part.getXmlAsync((result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded && typeof result.value === "string") {
        resolve(result.value);
      } else {
        reject(customXmlResultError(result, "无法读取公式源码元数据"));
      }
    });
  });
}

async function deleteCustomXmlPart(id: string): Promise<void> {
  try {
    const part = await getCustomXmlPart(id);
    await new Promise<void>((resolve, reject) => {
      part.deleteAsync((result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
        else reject(customXmlResultError(result, "无法删除旧公式源码元数据"));
      });
    });
  } catch (error) {
    console.warn("Formula metadata cleanup failed", error);
  }
}

function formulaTag(partId: string): string {
  return `${FORMULA_TAG_PREFIX}${partId}`;
}

function metadataPartId(tag: string): string | null {
  return tag.startsWith(FORMULA_TAG_PREFIX) ? tag.slice(FORMULA_TAG_PREFIX.length) : null;
}

function contentControlProperties(tag: string): string {
  return (
    `<w:sdtPr>` +
    `<w:alias w:val="Formula Canvas formula"/>` +
    `<w:tag w:val="${tag}"/>` +
    `</w:sdtPr>`
  );
}

function wrapFormulaContent(content: string, tag: string | null): string {
  return tag
    ? `<w:sdt>${contentControlProperties(tag)}<w:sdtContent>${content}</w:sdtContent></w:sdt>`
    : content;
}

function formulaMetadataXml(data: FormulaEditData): string {
  const payload = utf8ToBase64(JSON.stringify({ version: 1, ...data }));
  return (
    `<fc:formula xmlns:fc="${FORMULA_METADATA_NAMESPACE}" version="1">` +
    `<fc:data>${payload}</fc:data>` +
    `</fc:formula>`
  );
}

async function createFormulaMetadata(data: FormulaEditData): Promise<string | null> {
  try {
    const part = await addCustomXmlPart(formulaMetadataXml(data));
    return part.id || null;
  } catch (error) {
    // Formula insertion remains available on older Word builds, but those
    // formulas cannot participate in source round-tripping.
    console.warn("Formula metadata creation failed", error);
    return null;
  }
}

function setSelectedData(data: string, options: Office.SetSelectedDataOptions): Promise<void> {
  return new Promise((resolve, reject) => {
    Office.context.document.setSelectedDataAsync(data, options, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) resolve();
      else reject(new Error(result.error?.message || "Office 插入失败"));
    });
  });
}

export function isInsideOffice(): boolean {
  return typeof Office !== "undefined" && Boolean(Office.context?.document);
}

export function officeHostName(): string {
  if (!isInsideOffice()) return "浏览器预览";
  const host = Office.context.host;
  if (host === Office.HostType.Word) return "Word";
  if (host === Office.HostType.PowerPoint) return "PowerPoint";
  return "Office";
}

export function isWord(): boolean {
  return isInsideOffice() && Office.context.host === Office.HostType.Word;
}

export function supportsFormulaRoundTrip(): boolean {
  return (
    isWord() &&
    Office.context.requirements.isSetSupported("WordApi", "1.3") &&
    Office.context.requirements.isSetSupported("CustomXmlParts") &&
    Boolean(customXmlParts())
  );
}

/**
 * 把 OMML 片段包进最小的 WordprocessingML 段落，
 * 通过 Ooxml coercion 插入 Word，得到原生可编辑公式。
 *
 * 采用官方文档验证过的最小包结构：仅 `.rels` + `document.xml`。
 * `<m:oMath>` 直接放进段落内（行内数学区），Word 会渲染为可编辑公式。
 */
function buildOoxmlPackage(omml: string, tag: string | null, display: boolean): string {
  // mml2omml 输出的 <m:oMath> 自带 xmlns:m / xmlns:w 声明；
  // 这里剥掉它自带的命名空间声明，统一由 <w:document> 声明，避免重复/冲突。
  const oMathInner = omml
    .replace(/\sxmlns:m="[^"]*"/g, "")
    .replace(/\sxmlns:w="[^"]*"/g, "");

  const formula = display ? oMathInner : wrapFormulaContent(oMathInner, tag);
  const paragraph = `<w:p>${formula}</w:p>`;
  const bodyContent = display ? wrapFormulaContent(paragraph, tag) : paragraph;
  const documentXml =
    `<w:document ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:m="http://schemas.openxmlformats.org/officeDocument/2006/math">` +
    `<w:body>${bodyContent}</w:body>` +
    `</w:document>`;

  return (
    `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
    `<pkg:part pkg:name="/_rels/.rels" ` +
    `pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">` +
    `<pkg:xmlData>` +
    `<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" ` +
    `Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" ` +
    `Target="word/document.xml"/>` +
    `</Relationships>` +
    `</pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml" ` +
    `pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">` +
    `<pkg:xmlData>${documentXml}</pkg:xmlData>` +
    `</pkg:part>` +
    `</pkg:package>`
  );
}

function utf8ToBase64(value: string): string {
  const bytes = new TextEncoder().encode(value);
  let binary = "";
  const chunkSize = 0x8000;
  for (let offset = 0; offset < bytes.length; offset += chunkSize) {
    binary += String.fromCharCode(...bytes.subarray(offset, offset + chunkSize));
  }
  return btoa(binary);
}

const NUMBER_SLOT_WIDTH_PT = 42;
const DEFAULT_TEXT_WIDTH_PT = 468;

function formulaDrawingMarkup(
  widthEmu: number,
  heightEmu: number,
  baselinePositionHalfPoints = 0,
  includeSvg = true,
): string {
  const runProperties = baselinePositionHalfPoints
    ? `<w:rPr><w:position w:val="${baselinePositionHalfPoints}"/></w:rPr>`
    : "";
  const svgExtension = includeSvg
    ? `<a:extLst><a:ext uri="{96DAC541-7B7A-43D3-8B79-37D633B846F1}">` +
      `<asvg:svgBlip r:embed="rIdSvg"/>` +
      `</a:ext></a:extLst>`
    : "";
  const imageName = includeSvg ? "formula.svg" : "formula.png";
  return (
    `<w:r>${runProperties}<w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">` +
    `<wp:extent cx="${widthEmu}" cy="${heightEmu}"/>` +
    `<wp:effectExtent l="0" t="0" r="0" b="0"/>` +
    `<wp:docPr id="1" name="Formula Canvas equation"/>` +
    `<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect="1"/></wp:cNvGraphicFramePr>` +
    `<a:graphic><a:graphicData uri="http://schemas.openxmlformats.org/drawingml/2006/picture">` +
    `<pic:pic><pic:nvPicPr><pic:cNvPr id="0" name="${imageName}"/><pic:cNvPicPr/></pic:nvPicPr>` +
    `<pic:blipFill><a:blip r:embed="rIdPng">` +
    `${svgExtension}</a:blip><a:srcRect l="0" t="0" r="0" b="0"/>` +
    `<a:stretch><a:fillRect/></a:stretch></pic:blipFill>` +
    `<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="${widthEmu}" cy="${heightEmu}"/></a:xfrm>` +
    `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom><a:noFill/><a:ln><a:noFill/></a:ln></pic:spPr>` +
    `</pic:pic></a:graphicData></a:graphic>` +
    `</wp:inline></w:drawing></w:r>`
  );
}

function buildFormulaImagePackage(
  documentXml: string,
  rendered: RenderedFormula,
  pngBase64: string,
  includeSvg = true,
): string {
  const svgBase64 = utf8ToBase64(rendered.svg);
  const svgRelationship = includeSvg
    ? `<Relationship Id="rIdSvg" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/formula.svg"/>`
    : "";
  const svgPart = includeSvg
    ? `<pkg:part pkg:name="/word/media/formula.svg" pkg:contentType="image/svg+xml" pkg:compression="store">` +
      `<pkg:binaryData>${svgBase64}</pkg:binaryData></pkg:part>`
    : "";
  return (
    `<pkg:package xmlns:pkg="http://schemas.microsoft.com/office/2006/xmlPackage">` +
    `<pkg:part pkg:name="/_rels/.rels" ` +
    `pkg:contentType="application/vnd.openxmlformats-package.relationships+xml" pkg:padding="512">` +
    `<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>` +
    `</Relationships></pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/document.xml" ` +
    `pkg:contentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml">` +
    `<pkg:xmlData>${documentXml}</pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/_rels/document.xml.rels" ` +
    `pkg:contentType="application/vnd.openxmlformats-package.relationships+xml">` +
    `<pkg:xmlData><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">` +
    `<Relationship Id="rIdPng" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/formula.png"/>` +
    svgRelationship +
    `</Relationships></pkg:xmlData></pkg:part>` +
    `<pkg:part pkg:name="/word/media/formula.png" pkg:contentType="image/png" pkg:compression="store">` +
    `<pkg:binaryData>${pngBase64}</pkg:binaryData></pkg:part>` +
    svgPart +
    `</pkg:package>`
  );
}

/** Build an inline SVG run whose image baseline matches the surrounding text. */
export function buildInlineFormulaPackage(
  rendered: RenderedFormula,
  pngBase64: string,
  tag: string | null = null,
  includeSvg = true,
): string {
  const emuPerPoint = 12700;
  const widthEmu = Math.max(1, Math.round(rendered.widthPt * emuPerPoint));
  const heightEmu = Math.max(1, Math.round(rendered.heightPt * emuPerPoint));
  // w:position is a signed half-point value. A negative value lowers the run.
  const positionHalfPoints = -Math.round(rendered.baselineDepthPt * 2);
  const drawing = wrapFormulaContent(
    formulaDrawingMarkup(widthEmu, heightEmu, positionHalfPoints, includeSvg),
    tag,
  );
  const documentXml =
    `<w:document ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
    `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
    `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
    `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ` +
    `xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">` +
    `<w:body><w:p>${drawing}</w:p></w:body></w:document>`;
  return buildFormulaImagePackage(documentXml, rendered, pngBase64, includeSvg);
}

function buildDisplayFormulaPackage(
  rendered: RenderedFormula,
  pngBase64: string,
  tag: string | null,
  includeSvg: boolean,
): string {
  const emuPerPoint = 12700;
  const widthEmu = Math.max(1, Math.round(rendered.widthPt * emuPerPoint));
  const heightEmu = Math.max(1, Math.round(rendered.heightPt * emuPerPoint));
  const drawing = formulaDrawingMarkup(widthEmu, heightEmu, 0, includeSvg);
  const paragraph =
    `<w:p><w:pPr><w:jc w:val="center"/><w:spacing w:before="0" w:after="0"/></w:pPr>` +
    `${drawing}</w:p>`;
  const bodyContent = wrapFormulaContent(paragraph, tag);
  const documentXml =
    `<w:document ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
    `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
    `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
    `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ` +
    `xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">` +
    `<w:body>${bodyContent}</w:body></w:document>`;
  return buildFormulaImagePackage(documentXml, rendered, pngBase64, includeSvg);
}

/**
 * Build a MathType-style display-equation paragraph. A centered tab stop puts
 * the formula at the text-area midpoint and a right tab stop aligns the SEQ
 * field to the right margin. Keeping this as a normal paragraph avoids hidden
 * tables and preserves the requested point size without whole-image scaling.
 */
export function buildNumberedFormulaPackage(
  rendered: RenderedFormula,
  pngBase64: string,
  bookmarkName: string,
  textWidthPt: number | null = null,
  tag: string | null = null,
  numberFontSizePt = 12,
): string {
  const emuPerPoint = 12700;
  const widthEmu = Math.max(1, Math.round(rendered.widthPt * emuPerPoint));
  const heightEmu = Math.max(1, Math.round(rendered.heightPt * emuPerPoint));
  const drawing = formulaDrawingMarkup(widthEmu, heightEmu);
  const safeTextWidthPt =
    Number.isFinite(textWidthPt) && (textWidthPt as number) > 0
      ? Math.max(12, textWidthPt as number)
      : DEFAULT_TEXT_WIDTH_PT;
  const centerTabTwips = Math.round(safeTextWidthPt * 10);
  const rightTabTwips = Math.round(safeTextWidthPt * 20);
  const fontSizeHalfPoints = Math.max(1, Math.round(numberFontSizePt * 2));
  // Raise the number so it remains visually centered beside a multi-line SVG.
  const numberPositionHalfPoints = Math.max(0, Math.round(rendered.heightPt - numberFontSizePt));
  const numberRunProperties =
    `<w:rPr>` +
    (numberPositionHalfPoints
      ? `<w:position w:val="${numberPositionHalfPoints}"/>`
      : "") +
    `<w:sz w:val="${fontSizeHalfPoints}"/><w:szCs w:val="${fontSizeHalfPoints}"/>` +
    `</w:rPr>`;
  const paragraph =
    `<w:p><w:pPr><w:tabs><w:tab w:val="center" w:pos="${centerTabTwips}"/>` +
    `<w:tab w:val="right" w:pos="${rightTabTwips}"/></w:tabs>` +
    `<w:spacing w:before="0" w:after="0" w:lineRule="auto"/>` +
    `<w:ind w:left="0" w:right="0" w:firstLine="0"/></w:pPr>` +
    `<w:r><w:tab/></w:r>${drawing}<w:r><w:tab/></w:r>` +
    `<w:bookmarkStart w:id="0" w:name="${bookmarkName}"/>` +
    `<w:r>${numberRunProperties}<w:t>(</w:t></w:r>` +
    `<w:fldSimple w:instr=" SEQ FormulaCanvasEquation \\* ARABIC ">` +
    `<w:r>${numberRunProperties}<w:t>1</w:t></w:r></w:fldSimple>` +
    `<w:r>${numberRunProperties}<w:t>)</w:t></w:r><w:bookmarkEnd w:id="0"/>` +
    `</w:p>`;

  const documentXml =
    `<w:document ` +
    `xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
    `xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ` +
    `xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ` +
    `xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ` +
    `xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture" ` +
    `xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main">` +
    `<w:body>${wrapFormulaContent(paragraph, tag)}<w:p/></w:body></w:document>`;

  return buildFormulaImagePackage(documentXml, rendered, pngBase64);
}

async function refreshNumberField(bookmarkName: string): Promise<void> {
  if (
    typeof Word === "undefined" ||
    !Office.context.requirements.isSetSupported("WordApi", "1.5")
  ) {
    return;
  }

  try {
    await Word.run(async (context) => {
      const range = context.document.getBookmarkRangeOrNullObject(bookmarkName);
      range.load("isNullObject");
      await context.sync();
      if (range.isNullObject) return;

      const field = range.fields.getFirstOrNullObject();
      field.load("isNullObject");
      await context.sync();
      if (field.isNullObject) return;

      field.updateResult();
      await context.sync();
    });
  } catch (error) {
    // The SEQ field remains valid even on older Word builds; users can refresh
    // all fields with Cmd/Ctrl+A then F9.
    console.warn("Formula number field refresh failed", error);
  }
}

async function currentNumberedTextWidthPt(): Promise<number | null> {
  if (typeof Word === "undefined") {
    return null;
  }

  const supportsWordApi13 = Office.context.requirements.isSetSupported("WordApi", "1.3");
  const supportsDesktopLayout = Office.context.requirements.isSetSupported(
    "WordApiDesktop",
    "1.3",
  );
  if (!supportsWordApi13 && !supportsDesktopLayout) return null;

  try {
    return await Word.run(async (context) => {
      const selection = context.document.getSelection();
      const parentCell = supportsWordApi13 ? selection.parentTableCellOrNullObject : null;
      parentCell?.load("isNullObject,width");

      let setup: Word.PageSetup | null = null;
      let columns: Word.TextColumnCollection | null = null;
      if (supportsDesktopLayout) {
        const section = selection.sections.getFirst();
        setup = section.pageSetup;
        columns = setup.textColumns;
        setup.load("pageWidth,leftMargin,rightMargin,gutter");
        columns.load("items/width");
      }
      await context.sync();

      if (parentCell && !parentCell.isNullObject && Number.isFinite(parentCell.width)) {
        return Math.max(12, parentCell.width - 8);
      }
      if (!setup || !columns) return null;

      let width = setup.pageWidth - setup.leftMargin - setup.rightMargin - setup.gutter;
      const columnWidths = columns.items
        .map((column) => column.width)
        .filter((value) => Number.isFinite(value) && value > 0);
      if (columnWidths.length) width = Math.min(width, ...columnWidths);

      return Number.isFinite(width) && width > 0 ? Math.max(12, width) : null;
    });
  } catch (error) {
    console.warn("Unable to read the current Word text width", error);
    return null;
  }
}

async function selectFormulaForReplacement(
  target: FormulaReplacementTarget,
  newTag: string,
): Promise<void> {
  if (!supportsFormulaRoundTrip() || typeof Word === "undefined") {
    throw new Error("当前 Word 版本不支持替换已插入公式");
  }

  await Word.run(async (context) => {
    const control = context.document.contentControls.getByIdOrNullObject(target.controlId);
    control.load("isNullObject,tag");
    await context.sync();
    if (control.isNullObject) throw new Error("找不到原公式；它可能已被删除或取消组合");
    if (control.tag !== target.tag) {
      throw new Error("原公式已发生变化，请重新选中后再编辑");
    }
    // Retag the existing wrapper before replacement. Word may preserve the
    // selected wrapper or replace it with the wrapper in the OOXML package;
    // using the new tag in both cases keeps either outcome recoverable.
    control.tag = newTag;
    control.select("Select" as Word.SelectionMode);
    await context.sync();
  });
}

async function restoreFormulaTargetTag(
  target: FormulaReplacementTarget,
  currentTag: string,
): Promise<void> {
  if (typeof Word === "undefined") return;
  try {
    await Word.run(async (context) => {
      const control = context.document.contentControls.getByIdOrNullObject(target.controlId);
      control.load("isNullObject,tag");
      await context.sync();
      if (control.isNullObject || control.tag !== currentTag) return;
      control.tag = target.tag;
      await context.sync();
    });
  } catch (error) {
    console.warn("Formula content-control tag rollback failed", error);
  }
}

async function formulaControlExists(tag: string): Promise<boolean> {
  if (typeof Word === "undefined") return false;
  try {
    return await Word.run(async (context) => {
      const control = context.document.contentControls.getByTag(tag).getFirstOrNullObject();
      control.load("isNullObject");
      await context.sync();
      return !control.isNullObject;
    });
  } catch (error) {
    console.warn("Unable to check Formula Canvas content control", error);
    return true;
  }
}

/**
 * Word can preserve the selected content-control wrapper while inserting the
 * package's new wrapper, producing two nested controls with the same tag.
 * Remove only redundant nested wrappers and retain their formula contents.
 */
async function normalizeFormulaControls(tag: string): Promise<boolean> {
  if (typeof Word === "undefined") return false;
  try {
    return await Word.run(async (context) => {
      for (let pass = 0; pass < 8; pass += 1) {
        const controls = context.document.contentControls.getByTag(tag);
        controls.load("items/id,items/tag");
        await context.sync();
        if (!controls.items.length) return false;

        const parents = controls.items.map((control) => control.parentContentControlOrNullObject);
        parents.forEach((parent) => parent.load("isNullObject,tag"));
        await context.sync();

        const nestedIndex = parents.findIndex(
          (parent) => !parent.isNullObject && parent.tag === tag,
        );
        if (nestedIndex < 0) return true;
        controls.items[nestedIndex].delete(true);
        await context.sync();
      }
      // The formula remains linked even if a pathological document contains
      // more nesting than the defensive cleanup limit.
      return true;
    });
  } catch (error) {
    console.warn("Formula content-control normalization failed", error);
    return formulaControlExists(tag);
  }
}

function parseFormulaMetadata(xml: string): FormulaEditData {
  const documentXml = new DOMParser().parseFromString(xml, "application/xml");
  if (documentXml.querySelector("parsererror")) throw new Error("公式源码元数据已损坏");
  const dataNode = Array.from(documentXml.getElementsByTagNameNS("*", "data"))[0];
  if (!dataNode?.textContent) throw new Error("公式源码元数据为空");

  const parsed = JSON.parse(base64ToUtf8(dataNode.textContent.trim())) as Partial<FormulaEditData> & {
    version?: number;
  };
  const sizeRatios = parsed.sizeRatios ?? DEFAULT_MATHTYPE_SIZE_RATIOS;
  const validSizeRatios =
    typeof sizeRatios === "object" &&
    sizeRatios !== null &&
    [
      sizeRatios.subscript,
      sizeRatios.subSubscript,
      sizeRatios.symbol,
      sizeRatios.subSymbol,
    ].every((value) => typeof value === "number" && Number.isFinite(value) && value > 0);
  const preferences: InsertPreference[] = ["auto", "svg", "png", "omml", "numbered"];
  if (
    parsed.version !== 1 ||
    typeof parsed.source !== "string" ||
    typeof parsed.display !== "boolean" ||
    typeof parsed.color !== "string" ||
    typeof parsed.fontSizePt !== "number" ||
    !Number.isFinite(parsed.fontSizePt) ||
    !["transparent", "white"].includes(parsed.background ?? "") ||
    !["tex", "times-new-roman"].includes(parsed.mathFont ?? "") ||
    !["default", "upright", "italic", "bold", "bold-italic"].includes(parsed.mathStyle ?? "") ||
    !validSizeRatios ||
    !preferences.includes(parsed.preference as InsertPreference)
  ) {
    throw new Error("公式源码元数据格式不受支持");
  }
  return { ...parsed, sizeRatios } as FormulaEditData;
}

export async function recoverSelectedFormula(): Promise<RecoveredFormula> {
  if (!isWord()) throw new Error("重新编辑已插入公式仅支持 Word");
  if (!supportsFormulaRoundTrip() || typeof Word === "undefined") {
    throw new Error("当前 Word 版本不支持读取所选公式");
  }

  const target = await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const parent = selection.parentContentControlOrNullObject;
    const contained = selection.contentControls.getFirstOrNullObject();
    parent.load("isNullObject,tag,id");
    contained.load("isNullObject,tag,id");
    await context.sync();

    const candidates = [parent, contained];
    const formulaControl = candidates.find(
      (control) => !control.isNullObject && control.tag?.startsWith(FORMULA_TAG_PREFIX),
    );
    if (!formulaControl) {
      throw new Error("请在 Word 中选中由 Formula Canvas 新插入的公式后再试");
    }
    return { tag: formulaControl.tag, controlId: formulaControl.id };
  });

  const partId = metadataPartId(target.tag);
  if (!partId) throw new Error("所选内容不是 Formula Canvas 公式");
  const xml = await readCustomXml(await getCustomXmlPart(partId));
  return { ...parseFormulaMetadata(xml), ...target };
}

async function insertNumberedFormula(
  rendered: RenderedFormula,
  tag: string | null,
  fontSizePt: number,
): Promise<InsertResult> {
  const textWidthPt = await currentNumberedTextWidthPt();
  const effectiveTextWidthPt = textWidthPt ?? DEFAULT_TEXT_WIDTH_PT;
  if (rendered.widthPt > Math.max(1, effectiveTextWidthPt - NUMBER_SLOT_WIDTH_PT * 2 - 2)) {
    throw new Error(
      `公式宽度 ${rendered.widthPt.toFixed(1)} pt 超出编号段落的安全宽度；` +
        "请先使用 aligned / split 手动换行，插件不会缩小设定字号。",
    );
  }
  const pngBase64 = await svgToPngBase64(rendered, 4);
  const bookmarkName = `FormulaCanvasEq_${Date.now().toString(36)}`;
  const ooxml = buildNumberedFormulaPackage(
    rendered,
    pngBase64,
    bookmarkName,
    effectiveTextWidthPt,
    tag,
    fontSizePt,
  );
  try {
    await setSelectedData(ooxml, { coercionType: Office.CoercionType.Ooxml });
  } catch (e) {
    throw new Error(`编号公式插入失败：${e instanceof Error ? e.message : String(e)}`);
  }
  await refreshNumberField(bookmarkName);
  return { format: "NUMBERED", fallback: false, sourceLinked: Boolean(tag) };
}

async function insertBaselineAlignedPng(
  rendered: RenderedFormula,
  pngBase64?: string,
  tag: string | null = null,
): Promise<InsertResult> {
  if (
    typeof Word === "undefined" ||
    !Office.context.requirements.isSetSupported("WordApi", "1.3")
  ) {
    return insertPng(rendered);
  }

  const png = pngBase64 ?? (await svgToPngBase64(rendered, 4));
  await Word.run(async (context) => {
    const selection = context.document.getSelection();
    const picture = selection.insertInlinePictureFromBase64(
      png,
      Word.InsertLocation.replace,
    );
    picture.width = rendered.widthPt;
    picture.height = rendered.heightPt;
    picture.getRange().font.position = -rendered.baselineDepthPt;
    if (tag) {
      const control = picture.getRange().insertContentControl();
      control.tag = tag;
      control.title = "Formula Canvas formula";
    }
    await context.sync();
  });
  return { format: "PNG", fallback: false, sourceLinked: Boolean(tag) };
}

async function insertBaselineAlignedFormula(
  rendered: RenderedFormula,
  tag: string | null,
): Promise<InsertResult> {
  const pngBase64 = await svgToPngBase64(rendered, 4);
  const ooxml = buildInlineFormulaPackage(rendered, pngBase64, tag);
  try {
    await setSelectedData(ooxml, { coercionType: Office.CoercionType.Ooxml });
    return { format: "SVG", fallback: false, sourceLinked: Boolean(tag) };
  } catch (error) {
    console.warn("Baseline-aligned SVG insertion failed, falling back to PNG", error);
  }

  // WordApi 1.3 exposes the inserted picture's range, so even the compatibility
  // fallback can use the exact same point-based baseline correction.
  if (
    typeof Word !== "undefined" &&
    Office.context.requirements.isSetSupported("WordApi", "1.3")
  ) {
    const result = await insertBaselineAlignedPng(rendered, pngBase64, tag);
    return { ...result, fallback: true };
  }

  const result = await insertPng(rendered);
  return { ...result, fallback: true };
}

async function insertOmml(
  source: string,
  display: boolean,
  tag: string | null,
): Promise<InsertResult> {
  let mml: string;
  try {
    mml = renderMathML(source, display);
  } catch (e) {
    throw new Error(`MathML 生成失败：${e instanceof Error ? e.message : String(e)}`);
  }

  let omml: string;
  try {
    omml = mml2omml(mml);
  } catch (e) {
    throw new Error(`OMML 转换失败：${e instanceof Error ? e.message : String(e)}`);
  }
  if (!omml || !omml.includes("oMath")) {
    throw new Error(`OMML 结果异常：${(omml || "").slice(0, 80)}`);
  }

  const ooxml = buildOoxmlPackage(omml, tag, display);
  try {
    await setSelectedData(ooxml, { coercionType: Office.CoercionType.Ooxml });
  } catch (e) {
    throw new Error(`Word Ooxml 插入失败：${e instanceof Error ? e.message : String(e)}`);
  }
  return { format: "OMML", fallback: false, sourceLinked: Boolean(tag) };
}

async function insertPng(rendered: RenderedFormula): Promise<InsertResult> {
  const longestSide = Math.max(rendered.widthPx, rendered.heightPx);
  const scale = Math.max(1, Math.min(8, 4096 / longestSide));
  const png = await svgToPngBase64(rendered, scale);
  await setSelectedData(png, {
    coercionType: Office.CoercionType.Image,
    imageWidth: rendered.widthPt,
    imageHeight: rendered.heightPt,
  });
  return { format: "PNG", fallback: false, sourceLinked: false };
}

async function insertWordDisplayFormula(
  rendered: RenderedFormula,
  tag: string | null,
  preferSvg: boolean,
): Promise<InsertResult> {
  const pngBase64 = await svgToPngBase64(rendered, 4);
  if (preferSvg) {
    try {
      const svgPackage = buildDisplayFormulaPackage(rendered, pngBase64, tag, true);
      await setSelectedData(svgPackage, { coercionType: Office.CoercionType.Ooxml });
      return { format: "SVG", fallback: false, sourceLinked: Boolean(tag) };
    } catch (error) {
      console.warn("Display SVG insertion failed, falling back to packaged PNG", error);
    }
  }

  try {
    const pngPackage = buildDisplayFormulaPackage(rendered, pngBase64, tag, false);
    await setSelectedData(pngPackage, { coercionType: Office.CoercionType.Ooxml });
    return {
      format: "PNG",
      fallback: preferSvg,
      sourceLinked: Boolean(tag),
    };
  } catch (error) {
    console.warn("Packaged display PNG insertion failed, falling back to image coercion", error);
  }

  const result = await insertPng(rendered);
  return { ...result, fallback: preferSvg };
}

async function insertVectorFormula(rendered: RenderedFormula): Promise<InsertResult> {
  const imageOptions = {
    imageWidth: rendered.widthPt,
    imageHeight: rendered.heightPt,
  };

  const supportsSvg = Office.context.requirements.isSetSupported("ImageCoercion", "1.2");
  if (supportsSvg) {
    try {
      // XmlSvg expects the SVG XML string itself. Unlike CoercionType.Image,
      // passing Base64 here makes Office 2021 store the Base64 text verbatim.
      await setSelectedData(rendered.svg, {
        coercionType: Office.CoercionType.XmlSvg,
        ...imageOptions,
      });
      return { format: "SVG", fallback: false, sourceLinked: false };
    } catch (error) {
      console.warn("SVG insertion failed, falling back to PNG", error);
    }
  }

  const result = await insertPng(rendered);
  return { ...result, fallback: true };
}

export async function insertFormula(
  rendered: RenderedFormula,
  preference: InsertPreference = "auto",
  math?: Omit<FormulaEditData, "preference">,
  replacementTarget: FormulaReplacementTarget | null = null,
): Promise<InsertResult> {
  if (!isInsideOffice()) throw new Error("请在 Word 或 PowerPoint 中打开插件后插入");

  let newPartId: string | null = null;
  let tag: string | null = null;
  if (isWord() && math) {
    if (supportsFormulaRoundTrip()) {
      newPartId = await createFormulaMetadata({ ...math, preference });
      tag = newPartId ? formulaTag(newPartId) : null;
    }
    if (replacementTarget && !tag) {
      throw new Error("无法保存新的公式源码；已保留 Word 中的原公式");
    }
  }

  try {
    if (replacementTarget && tag) {
      await selectFormulaForReplacement(replacementTarget, tag);
    }

    let result: InsertResult;

    if (preference === "omml") {
      if (!isWord()) throw new Error("可编辑公式仅支持 Word");
      if (!math) throw new Error("缺少公式源码，无法生成可编辑公式");
      try {
        result = await insertOmml(math.source, math.display, tag);
      } catch (error) {
        console.warn("OMML insertion failed, falling back to SVG/PNG", error);
        // 行内 OMML 回退仍必须使用相同的基线修正，避免下标公式上浮。
        const fallbackResult = math.display
          ? await insertWordDisplayFormula(rendered, tag, true)
          : await insertBaselineAlignedFormula(rendered, tag);
        result = { ...fallbackResult, fallback: true };
      }
    } else if (isWord() && math && !math.display) {
      // Word otherwise aligns the bottom of an inline image to the text baseline,
      // lifting formulas that extend below their own math baseline.
      result =
        preference === "png"
          ? await insertBaselineAlignedPng(rendered, undefined, tag)
          : await insertBaselineAlignedFormula(rendered, tag);
    } else if (preference === "numbered") {
      if (!isWord()) throw new Error("自动编号公式仅支持 Word");
      if (!math) throw new Error("缺少排版信息，无法判断行内或独立公式");
      result = await insertNumberedFormula(rendered, tag, math.fontSizePt);
    } else if (isWord() && math) {
      result = await insertWordDisplayFormula(rendered, tag, preference !== "png");
    } else if (preference === "png") {
      result = await insertPng(rendered);
    } else {
      result = await insertVectorFormula(rendered);
    }

    if (tag && isWord()) {
      result = { ...result, sourceLinked: await normalizeFormulaControls(tag) };
    }
    if (newPartId && !result.sourceLinked) await deleteCustomXmlPart(newPartId);
    if (replacementTarget) {
      const oldPartId = metadataPartId(replacementTarget.tag);
      if (oldPartId && !(await formulaControlExists(replacementTarget.tag))) {
        await deleteCustomXmlPart(oldPartId);
      }
    }
    return result;
  } catch (error) {
    if (replacementTarget && tag) {
      await restoreFormulaTargetTag(replacementTarget, tag);
    }
    if (newPartId && (!tag || !(await formulaControlExists(tag)))) {
      await deleteCustomXmlPart(newPartId);
    }
    throw error;
  }
}
