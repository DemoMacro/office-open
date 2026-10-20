import { createColorElement } from "@office-open/core/drawing";
import type { SolidFillOptions } from "@office-open/core/drawing";
import type { ShowOptions } from "@shared/file";

/** Web output options (p:webPr — HTML export behavior). */
export interface WebPropertiesOptions {
  showAnimation?: boolean;
  resizeGraphics?: boolean;
  allowPng?: boolean;
  relyOnVml?: boolean;
  organizeInFolders?: boolean;
  useLongFilenames?: boolean;
  /** Monitor size for web output (p:webPr `@imgSz`, ST_WebScreenSize) */
  imageSize?:
    | "544x376"
    | "640x480"
    | "720x512"
    | "800x600"
    | "1024x768"
    | "1152x882"
    | "1152x900"
    | "1280x1024"
    | "1600x1200"
    | "1800x1400"
    | "1920x1200";
  encoding?: string;
  /** Color scheme for web output (p:webPr `@clr`, ST_WebColorType) */
  color?:
    | "none"
    | "browser"
    | "presentationText"
    | "presentationAccent"
    | "whiteTextOnBlack"
    | "blackTextOnWhite";
}

/** Print options (p:prnPr — what each page holds and the color mode). */
export interface PrintPropertiesOptions {
  /** What each printed page holds: slides, "handoutsN" N slides per page, notes pages, or the outline. */
  printWhat?:
    | "slides"
    | "handouts1"
    | "handouts2"
    | "handouts3"
    | "handouts4"
    | "handouts6"
    | "handouts9"
    | "notes"
    | "outline";
  /** Print in pure black-and-white, grayscale, or color. */
  colorMode?: "blackWhite" | "gray" | "color";
  hiddenSlides?: boolean;
  scaleToFitPaper?: boolean;
  frameSlides?: boolean;
}

/** Web-publishing target (p:htmlPubPr — browser and title for HTML export). */
export type HtmlPublishSlideScopeOptions =
  | { kind: "all" }
  | { kind: "range"; start: number; end: number }
  | { kind: "customShow"; id: number };

export interface HtmlPublishPropertiesOptions {
  showSpeakerNotes?: boolean;
  target?: string;
  title?: string;
  /** Publishing relationship target (`r:id`, required by CT_HtmlPublishProperties). */
  rId: string;
  /** Required EG_SlideListChoice child: all slides, an index range, or a custom show. */
  slideScope: HtmlPublishSlideScopeOptions;
  /** Verbatim inner XML of p:htmlPubPr/p:extLst. */
  ext?: string;
}

export interface PresentationPropertiesOptions {
  show?: ShowOptions;
  web?: WebPropertiesOptions;
  print?: PrintPropertiesOptions;
  htmlPublish?: HtmlPublishPropertiesOptions;
  colorMru?: string[];
  /** Verbatim inner XML of p:extLst (last child per CT_PresentationProperties). */
  ext?: string;
}

const COLOR_MODE_XSD: Record<string, string> = {
  blackWhite: "bw",
  gray: "gray",
  color: "clr",
};

function buildWebPrXml(opts: WebPropertiesOptions): string {
  const attrs: string[] = [];
  if (opts.showAnimation) attrs.push(' showAnimation="1"');
  if (opts.resizeGraphics === false) attrs.push(' resizeGraphics="0"');
  if (opts.allowPng) attrs.push(' allowPng="1"');
  if (opts.relyOnVml) attrs.push(' relyOnVml="1"');
  if (opts.organizeInFolders === false) attrs.push(' organizeInFolders="0"');
  if (opts.useLongFilenames === false) attrs.push(' useLongFilenames="0"');
  if (opts.imageSize) attrs.push(` imgSz="${opts.imageSize}"`);
  if (opts.encoding) attrs.push(` encoding="${opts.encoding}"`);
  if (opts.color) attrs.push(` clr="${opts.color}"`);
  return `<p:webPr${attrs.join("")}/>`;
}

function buildPrnPrXml(opts: PrintPropertiesOptions): string {
  const attrs: string[] = [];
  if (opts.printWhat) attrs.push(` prnWhat="${opts.printWhat}"`);
  if (opts.colorMode) attrs.push(` clrMode="${COLOR_MODE_XSD[opts.colorMode]}"`);
  if (opts.hiddenSlides) attrs.push(' hiddenSlides="1"');
  if (opts.scaleToFitPaper) attrs.push(' scaleToFitPaper="1"');
  if (opts.frameSlides) attrs.push(' frameSlides="1"');
  return `<p:prnPr${attrs.join("")}/>`;
}

function buildHtmlPubPrXml(opts: HtmlPublishPropertiesOptions): string {
  const attrs: string[] = [];
  if (opts.showSpeakerNotes === false) attrs.push(' showSpeakerNotes="0"');
  if (opts.target) attrs.push(` target="${opts.target}"`);
  if (opts.title) attrs.push(` title="${opts.title}"`);
  attrs.push(` r:id="${opts.rId}"`);
  const scope = opts.slideScope;
  const scopeXml =
    scope.kind === "all"
      ? "<p:sldAll/>"
      : scope.kind === "range"
        ? `<p:sldRg st="${scope.start}" end="${scope.end}"/>`
        : `<p:custShow id="${scope.id}"/>`;
  const extensionXml = opts.ext === undefined ? "" : `<p:extLst>${opts.ext}</p:extLst>`;
  return `<p:htmlPubPr${attrs.join("")}>${scopeXml}${extensionXml}</p:htmlPubPr>`;
}

function buildShowPrXml(showOptions: ShowOptions): string {
  const showPrAttrs: string[] = [];
  if (showOptions.loop !== undefined) showPrAttrs.push(` loop="${showOptions.loop ? 1 : 0}"`);
  if (showOptions.showNarration !== undefined)
    showPrAttrs.push(` showNarration="${showOptions.showNarration ? 1 : 0}"`);
  if (showOptions.showAnimation !== undefined)
    showPrAttrs.push(` showAnimation="${showOptions.showAnimation ? 1 : 0}"`);
  if (showOptions.useTimings !== undefined)
    showPrAttrs.push(` useTimings="${showOptions.useTimings ? 1 : 0}"`);

  const sourceType = showOptions.type;
  const showType = sourceType === null ? undefined : (sourceType ?? "present");
  let showTypeXml = "";
  if (showType === "kiosk") {
    const restartAttr =
      showOptions.restart !== undefined ? ` restart="${showOptions.restart}"` : "";
    showTypeXml = `<p:kiosk${restartAttr}/>`;
  } else if (showType === "browse") {
    const scrollbarAttr = showOptions.showScrollbar === false ? ' showScrollbar="0"' : "";
    showTypeXml = `<p:browse${scrollbarAttr}/>`;
  } else if (showType !== undefined) {
    showTypeXml = "<p:present/>";
  }

  let slideListXml = showOptions.slideSelection === null ? "" : "<p:sldAll/>";
  if (showOptions.slideRange) {
    slideListXml = `<p:sldRg st="${showOptions.slideRange.start}" end="${showOptions.slideRange.end}"/>`;
  } else if (showOptions.customShowId !== undefined) {
    slideListXml = `<p:custShow id="${showOptions.customShowId}"/>`;
  }

  let penClrXml = "";
  if (showOptions.penColor) {
    const color: SolidFillOptions =
      typeof showOptions.penColor === "string"
        ? { value: showOptions.penColor }
        : showOptions.penColor;
    penClrXml = `<p:penClr>${createColorElement(color)}</p:penClr>`;
  }

  const extXml = showOptions.ext !== undefined ? `<p:extLst>${showOptions.ext}</p:extLst>` : "";
  return `<p:showPr${showPrAttrs.join("")}>${showTypeXml}${slideListXml}${penClrXml}${extXml}</p:showPr>`;
}

export function buildPresentationPropertiesXml(opts?: PresentationPropertiesOptions): string {
  const ns =
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

  if (!opts) {
    return `<p:presentationPr ${ns}/>`;
  }

  const children: string[] = [];
  if (opts.htmlPublish) children.push(buildHtmlPubPrXml(opts.htmlPublish));
  if (opts.web) children.push(buildWebPrXml(opts.web));
  if (opts.print) children.push(buildPrnPrXml(opts.print));
  if (opts.show) children.push(buildShowPrXml(opts.show));
  if (opts.colorMru && opts.colorMru.length > 0) {
    children.push(
      `<p:clrMru>${opts.colorMru.map((c) => `<a:srgbClr val="${c}"/>`).join("")}</p:clrMru>`,
    );
  }
  if (opts.ext) children.push(`<p:extLst>${opts.ext}</p:extLst>`);

  if (children.length === 0) {
    return `<p:presentationPr ${ns}/>`;
  }

  return `<p:presentationPr ${ns}>${children.join("")}</p:presentationPr>`;
}
