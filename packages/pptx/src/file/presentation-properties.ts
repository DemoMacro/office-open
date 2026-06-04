import type { Context } from "@file/xml-components";
import { ImportedXmlComponent } from "@file/xml-components";

import type { ShowOptions } from "./file";

export interface WebPropertiesOptions {
  readonly showAnimation?: boolean;
  readonly resizeGraphics?: boolean;
  readonly allowPng?: boolean;
  readonly relyOnVml?: boolean;
  readonly organizeInFolders?: boolean;
  readonly useLongFilenames?: boolean;
  readonly imageSize?: string;
  readonly encoding?: string;
  readonly color?: string;
}

export interface PrintPropertiesOptions {
  readonly printWhat?:
    | "slides"
    | "handouts1"
    | "handouts2"
    | "handouts3"
    | "handouts4"
    | "handouts6"
    | "handouts9"
    | "notes"
    | "outline";
  readonly colorMode?: "blackWhite" | "gray" | "color";
  readonly hiddenSlides?: boolean;
  readonly scaleToFitPaper?: boolean;
  readonly frameSlides?: boolean;
}

export interface HtmlPublishPropertiesOptions {
  readonly showSpeakerNotes?: boolean;
  readonly target?: string;
  readonly title?: string;
  readonly rId?: string;
}

export interface PresentationPropertiesFullOptions {
  readonly show?: ShowOptions;
  readonly web?: WebPropertiesOptions;
  readonly print?: PrintPropertiesOptions;
  readonly htmlPublish?: HtmlPublishPropertiesOptions;
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
  if (opts.rId) attrs.push(` r:id="${opts.rId}"`);
  return `<p:htmlPubPr${attrs.join("")}><p:sldAll/></p:htmlPubPr>`;
}

function buildShowPrXml(showOptions: ShowOptions): string {
  const showPrAttrs: string[] = [];
  if (showOptions.loop) showPrAttrs.push(' loop="1"');
  if (showOptions.showNarration === false) showPrAttrs.push(' showNarration="0"');
  if (showOptions.showAnimation === false) showPrAttrs.push(' showAnimation="0"');
  if (showOptions.useTimings) showPrAttrs.push(' useTimings="1"');

  const showType = showOptions.type ?? "present";
  let showTypeXml: string;
  if (showType === "kiosk") {
    const restartAttr =
      showOptions.restart !== undefined ? ` restart="${showOptions.restart}"` : "";
    showTypeXml = `<p:kiosk${restartAttr}/>`;
  } else if (showType === "browse") {
    const scrollbarAttr = showOptions.showScrollbar === false ? ' showScrollbar="0"' : "";
    showTypeXml = `<p:browse${scrollbarAttr}/>`;
  } else {
    showTypeXml = "<p:present/>";
  }

  let slideListXml = "<p:sldAll/>";
  if (showOptions.slideRange) {
    slideListXml = `<p:sldRg st="${showOptions.slideRange.start}" end="${showOptions.slideRange.end}"/>`;
  }

  let penClrXml = "";
  if (showOptions.penColor) {
    penClrXml = `<p:penClr><a:srgbClr val="${showOptions.penColor}"/></p:penClr>`;
  }

  return `<p:showPr${showPrAttrs.join("")}>${showTypeXml}${slideListXml}${penClrXml}</p:showPr>`;
}

function buildPresPropsXml(opts?: PresentationPropertiesFullOptions): string {
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

  if (children.length === 0) {
    return `<p:presentationPr ${ns}/>`;
  }

  return `<p:presentationPr ${ns}>${children.join("")}</p:presentationPr>`;
}

export class PresentationProperties extends ImportedXmlComponent {
  private static cache = new Map<string, ImportedXmlComponent>();
  private readonly key: string;

  public constructor(opts?: PresentationPropertiesFullOptions) {
    super("p:presentationPr");
    this.key = opts ? JSON.stringify(opts) : "";
    if (!PresentationProperties.cache.has(this.key)) {
      PresentationProperties.cache.set(
        this.key,
        ImportedXmlComponent.fromXmlString(buildPresPropsXml(opts)),
      );
    }
  }

  public override toXml(context: Context): string {
    return PresentationProperties.cache.get(this.key)!.toXml(context);
  }
}
