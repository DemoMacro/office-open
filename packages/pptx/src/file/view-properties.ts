import type { Context } from "@file/xml-components";
import { ImportedXmlComponent } from "@file/xml-components";

export interface ViewPropertiesOptions {
  /** Last active view type */
  readonly lastView?:
    | "sldView"
    | "sldMasterView"
    | "notesView"
    | "handoutView"
    | "outlineView"
    | "sldSorterView";
  /** Show comments (default true) */
  readonly showComments?: boolean;
  /** Grid spacing (EMU, default 72008) */
  readonly gridSpacing?: { readonly cx: number; readonly cy: number };
  /** Zoom scale numerator (default 90) */
  readonly zoomScaleNumerator?: number;
  /** Zoom scale denominator (default 100) */
  readonly zoomScaleDenominator?: number;
}

function buildViewPropsXml(opts?: ViewPropertiesOptions): string {
  const ns =
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

  if (!opts) {
    // Default view properties (matches Excel/PowerPoint defaults)
    return [
      `<p:viewPr ${ns}>`,
      "<p:normalViewPr>",
      '<p:restoredLeft sz="14996" autoAdjust="0"/>',
      '<p:restoredTop sz="94660"/>',
      "</p:normalViewPr>",
      "<p:slideViewPr>",
      '<p:cSldViewPr snapToGrid="0">',
      '<p:cViewPr varScale="1">',
      '<p:scale><a:sx n="90" d="100"/><a:sy n="90" d="100"/></p:scale>',
      '<p:origin x="1200" y="72"/>',
      "</p:cViewPr>",
      "<p:guideLst/>",
      "</p:cSldViewPr>",
      "</p:slideViewPr>",
      "<p:notesTextViewPr>",
      "<p:cViewPr>",
      '<p:scale><a:sx n="1" d="1"/><a:sy n="1" d="1"/></p:scale>',
      '<p:origin x="0" y="0"/>',
      "</p:cViewPr>",
      "</p:notesTextViewPr>",
      '<p:gridSpacing cx="72008" cy="72008"/>',
      "</p:viewPr>",
    ].join("");
  }

  const zoomN = opts.zoomScaleNumerator ?? 90;
  const zoomD = opts.zoomScaleDenominator ?? 100;
  const gridCx = opts.gridSpacing?.cx ?? 72008;
  const gridCy = opts.gridSpacing?.cy ?? 72008;

  const parts: string[] = [];

  // Root attributes
  let rootAttrs = "";
  if (opts.lastView) rootAttrs += ` lastView="${opts.lastView}"`;
  if (opts.showComments === false) rootAttrs += ' showComments="0"';

  parts.push(`<p:viewPr ${ns}${rootAttrs}>`);

  // normalViewPr
  parts.push("<p:normalViewPr>");
  parts.push('<p:restoredLeft sz="14996" autoAdjust="0"/>');
  parts.push('<p:restoredTop sz="94660"/>');
  parts.push("</p:normalViewPr>");

  // slideViewPr
  parts.push("<p:slideViewPr>");
  parts.push('<p:cSldViewPr snapToGrid="0">');
  parts.push('<p:cViewPr varScale="1">');
  parts.push(
    `<p:scale><a:sx n="${zoomN}" d="${zoomD}"/><a:sy n="${zoomN}" d="${zoomD}"/></p:scale>`,
  );
  parts.push('<p:origin x="1200" y="72"/>');
  parts.push("</p:cViewPr>");
  parts.push("<p:guideLst/>");
  parts.push("</p:cSldViewPr>");
  parts.push("</p:slideViewPr>");

  // notesTextViewPr
  parts.push("<p:notesTextViewPr>");
  parts.push("<p:cViewPr>");
  parts.push('<p:scale><a:sx n="1" d="1"/><a:sy n="1" d="1"/></p:scale>');
  parts.push('<p:origin x="0" y="0"/>');
  parts.push("</p:cViewPr>");
  parts.push("</p:notesTextViewPr>");

  // gridSpacing
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
