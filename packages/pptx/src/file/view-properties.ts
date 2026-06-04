import type { Context } from "@file/xml-components";
import { ImportedXmlComponent } from "@file/xml-components";

type SplitterBarState = "restored" | "maximized" | "minimized";

export interface NormalViewOptions {
  readonly showOutlineIcons?: boolean;
  readonly snapVertSplitter?: boolean;
  readonly vertBarState?: SplitterBarState;
  readonly horzBarState?: SplitterBarState;
  readonly preferSingleView?: boolean;
}

export interface SlideViewOptions {
  readonly snapToGrid?: boolean;
  readonly snapToObjects?: boolean;
  readonly showGuides?: boolean;
  readonly varScale?: boolean;
}

const LAST_VIEW_XSD: Record<string, string> = {
  slideView: "sldView",
  slideMasterView: "sldMasterView",
  notesView: "notesView",
  handoutView: "handoutView",
  outlineView: "outlineView",
  slideSorterView: "sldSorterView",
};

export interface ViewPropertiesOptions {
  readonly lastView?:
    | "slideView"
    | "slideMasterView"
    | "notesView"
    | "handoutView"
    | "outlineView"
    | "slideSorterView";
  readonly showComments?: boolean;
  readonly gridSpacing?: { readonly cx: number; readonly cy: number };
  readonly zoomScaleNumerator?: number;
  readonly zoomScaleDenominator?: number;
  readonly normalView?: NormalViewOptions;
  readonly slideView?: SlideViewOptions;
}

function buildNormalViewPrXml(opts?: NormalViewOptions): string {
  const attrs: string[] = [];
  if (opts?.showOutlineIcons === false) attrs.push(' showOutlineIcons="0"');
  if (opts?.snapVertSplitter) attrs.push(' snapVertSplitter="1"');
  if (opts?.vertBarState) attrs.push(` vertBarState="${opts.vertBarState}"`);
  if (opts?.horzBarState) attrs.push(` horzBarState="${opts.horzBarState}"`);
  if (opts?.preferSingleView) attrs.push(' preferSingleView="1"');
  return (
    `<p:normalViewPr${attrs.join("")}>` +
    '<p:restoredLeft sz="14996" autoAdjust="0"/>' +
    '<p:restoredTop sz="94660"/>' +
    "</p:normalViewPr>"
  );
}

function buildCSldViewPrXml(opts?: SlideViewOptions, zoomN = 90, zoomD = 100): string {
  const attrs: string[] = [];
  if (opts?.snapToGrid === false) attrs.push(' snapToGrid="0"');
  if (opts?.snapToObjects) attrs.push(' snapToObjects="1"');
  if (opts?.showGuides) attrs.push(' showGuides="1"');

  const varScale = opts?.varScale !== false;
  const cViewPrAttr = varScale ? ' varScale="1"' : "";

  return (
    `<p:cSldViewPr${attrs.join("")}>` +
    `<p:cViewPr${cViewPrAttr}>` +
    `<p:scale><a:sx n="${zoomN}" d="${zoomD}"/><a:sy n="${zoomN}" d="${zoomD}"/></p:scale>` +
    '<p:origin x="1200" y="72"/>' +
    "</p:cViewPr>" +
    "<p:guideLst/>" +
    "</p:cSldViewPr>"
  );
}

function buildViewPropsXml(opts?: ViewPropertiesOptions): string {
  const ns =
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

  const zoomN = opts?.zoomScaleNumerator ?? 90;
  const zoomD = opts?.zoomScaleDenominator ?? 100;
  const gridCx = opts?.gridSpacing?.cx ?? 72008;
  const gridCy = opts?.gridSpacing?.cy ?? 72008;

  const parts: string[] = [];

  let rootAttrs = "";
  if (opts?.lastView) rootAttrs += ` lastView="${LAST_VIEW_XSD[opts.lastView]}"`;
  if (opts?.showComments === false) rootAttrs += ' showComments="0"';

  parts.push(`<p:viewPr ${ns}${rootAttrs}>`);
  parts.push(buildNormalViewPrXml(opts?.normalView));
  parts.push("<p:slideViewPr>");
  parts.push(buildCSldViewPrXml(opts?.slideView, zoomN, zoomD));
  parts.push("</p:slideViewPr>");
  parts.push("<p:notesTextViewPr>");
  parts.push("<p:cViewPr>");
  parts.push('<p:scale><a:sx n="1" d="1"/><a:sy n="1" d="1"/></p:scale>');
  parts.push('<p:origin x="0" y="0"/>');
  parts.push("</p:cViewPr>");
  parts.push("</p:notesTextViewPr>");
  parts.push(`<p:gridSpacing cx="${gridCx}" cy="${gridCy}"/>`);
  parts.push("</p:viewPr>");
  return parts.join("");
}

export class ViewProperties extends ImportedXmlComponent {
  private static cache = new Map<string, ImportedXmlComponent>();
  private readonly key: string;

  public constructor(opts?: ViewPropertiesOptions) {
    super("p:viewPr");
    this.key = opts ? JSON.stringify(opts) : "";
    if (!ViewProperties.cache.has(this.key)) {
      ViewProperties.cache.set(
        this.key,
        ImportedXmlComponent.fromXmlString(buildViewPropsXml(opts)),
      );
    }
  }

  public override toXml(context: Context): string {
    return ViewProperties.cache.get(this.key)!.toXml(context);
  }
}
