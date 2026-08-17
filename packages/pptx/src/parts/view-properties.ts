type SplitterBarState = "restored" | "maximized" | "minimized";

/**
 * CT_CommonViewProperties (`p:cViewPr`) — zoom scale and origin, the child
 * shared by the slide/outline/notes-text/sorter view property elements.
 */
export interface CommonViewPropertiesOptions {
  /** Zoom scale as a fraction (CT_Scale2D `a:sx` n/d; `a:sy` mirrors sx). */
  scale?: { numerator: number; denominator: number };
  /** View origin coordinates (CT_Point2D). */
  origin?: { x: number; y: number };
  /** Variable scaling allowed (`@varScale`, default false). */
  variableScale?: boolean;
}

/** CT_NormalViewPortion — one restored pane of the normal view. */
export interface NormalViewPortionOptions {
  /** Restored pane size in thousandths of a percent (`@sz`, required). */
  size: number;
  /** Auto-adjust the pane when the window resizes (`@autoAdjust`, default true). */
  autoAdjust?: boolean;
}

export interface NormalViewOptions {
  showOutlineIcons?: boolean;
  snapVertSplitter?: boolean;
  vertBarState?: SplitterBarState;
  horzBarState?: SplitterBarState;
  preferSingleView?: boolean;
  /** Restored left outline pane (`p:restoredLeft`). */
  restoredLeft?: NormalViewPortionOptions;
  /** Restored top notes pane (`p:restoredTop`). */
  restoredTop?: NormalViewPortionOptions;
}

export interface SlideViewOptions {
  snapToGrid?: boolean;
  snapToObjects?: boolean;
  showGuides?: boolean;
  /** Zoom/origin of the slide view (`p:cViewPr`). */
  view?: CommonViewPropertiesOptions;
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
  lastView?:
    | "slideView"
    | "slideMasterView"
    | "notesView"
    | "handoutView"
    | "outlineView"
    | "slideSorterView";
  showComments?: boolean;
  gridSpacing?: { cx: number; cy: number };
  normalView?: NormalViewOptions;
  slideView?: SlideViewOptions;
  guides?: {
    orient?: "vert" | "horz";
    pos?: number;
  }[];
  /** Outline view (`p:outlineViewPr`) — view is its required `p:cViewPr`. */
  outlineView?: {
    view: CommonViewPropertiesOptions;
    slides?: {
      rId: string;
      collapse?: boolean;
    }[];
  };
  /** Notes text view (`p:notesTextViewPr`) — the `p:cViewPr` payload. */
  notesTextView?: CommonViewPropertiesOptions;
  sorterView?: {
    showFormatting?: boolean;
    view: CommonViewPropertiesOptions;
  };
  notesView?: boolean;
}

function portionXml(
  name: string,
  opts: NormalViewPortionOptions | undefined,
  fallbackSize: number,
): string {
  const size = opts?.size ?? fallbackSize;
  const attrs = [`sz="${size}"`];
  if (opts?.autoAdjust !== undefined) attrs.push(`autoAdjust="${opts.autoAdjust ? 1 : 0}"`);
  return `<p:${name} ${attrs.join(" ")}/>`;
}

function commonViewPrXml(opts: CommonViewPropertiesOptions | undefined): string {
  // scale and origin are required children; defaults match a fresh Office file.
  const n = opts?.scale?.numerator ?? 90;
  const d = opts?.scale?.denominator ?? 100;
  const x = opts?.origin?.x ?? 1200;
  const y = opts?.origin?.y ?? 72;
  const varScale = opts?.variableScale ? ' varScale="1"' : "";
  return (
    `<p:cViewPr${varScale}>` +
    `<p:scale><a:sx n="${n}" d="${d}"/><a:sy n="${n}" d="${d}"/></p:scale>` +
    `<p:origin x="${x}" y="${y}"/>` +
    "</p:cViewPr>"
  );
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
    portionXml("restoredLeft", opts?.restoredLeft, 15619) +
    portionXml("restoredTop", opts?.restoredTop, 94681) +
    "</p:normalViewPr>"
  );
}

function buildCSldViewPrXml(
  opts?: SlideViewOptions,
  guides?: ViewPropertiesOptions["guides"],
): string {
  const attrs: string[] = [];
  if (opts?.snapToGrid === false) attrs.push(' snapToGrid="0"');
  if (opts?.snapToObjects) attrs.push(' snapToObjects="1"');
  if (opts?.showGuides) attrs.push(' showGuides="1"');

  // Build guideLst with optional guides
  let guideLstXml = "<p:guideLst/>";
  if (guides && guides.length > 0) {
    const guideXmls = guides.map((g) => {
      const gAttrs: string[] = [];
      if (g.orient) gAttrs.push(` orient="${g.orient}"`);
      if (g.pos !== undefined) gAttrs.push(` pos="${g.pos}"`);
      return `<p:guide${gAttrs.join("")}/>`;
    });
    guideLstXml = `<p:guideLst>${guideXmls.join("")}</p:guideLst>`;
  }

  return (
    `<p:cSldViewPr${attrs.join("")}>` +
    commonViewPrXml(opts?.view) +
    guideLstXml +
    "</p:cSldViewPr>"
  );
}

export function buildViewPropsXml(opts?: ViewPropertiesOptions): string {
  const ns =
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

  const gridCx = opts?.gridSpacing?.cx ?? 72008;
  const gridCy = opts?.gridSpacing?.cy ?? 72008;

  const parts: string[] = [];

  let rootAttrs = "";
  if (opts?.lastView) rootAttrs += ` lastView="${LAST_VIEW_XSD[opts.lastView]}"`;
  if (opts?.showComments === false) rootAttrs += ' showComments="0"';

  parts.push(`<p:viewPr ${ns}${rootAttrs}>`);
  parts.push(buildNormalViewPrXml(opts?.normalView));
  parts.push("<p:slideViewPr>");
  parts.push(buildCSldViewPrXml(opts?.slideView, opts?.guides));
  parts.push("</p:slideViewPr>");

  // outlineViewPr — sequence: cViewPr (required), sldLst
  if (opts?.outlineView) {
    let outlineXml = "<p:outlineViewPr>";
    outlineXml += commonViewPrXml(opts.outlineView.view);
    if (opts.outlineView.slides && opts.outlineView.slides.length > 0) {
      outlineXml += "<p:sldLst>";
      for (const sl of opts.outlineView.slides) {
        const slAttrs: string[] = [`r:id="${sl.rId}"`];
        if (sl.collapse) slAttrs.push(' collapse="1"');
        outlineXml += `<p:sld ${slAttrs.join(" ")}/>`;
      }
      outlineXml += "</p:sldLst>";
    }
    outlineXml += "</p:outlineViewPr>";
    parts.push(outlineXml);
  }

  // notesTextViewPr — cViewPr only
  if (opts?.notesTextView) {
    parts.push(`<p:notesTextViewPr>${commonViewPrXml(opts.notesTextView)}</p:notesTextViewPr>`);
  }

  // sorterViewPr — sequence: cViewPr (not cSldViewPr!)
  if (opts?.sorterView) {
    const sorterAttrs: string[] = [];
    if (opts.sorterView.showFormatting) sorterAttrs.push(' showFormatting="1"');
    parts.push(`<p:sorterViewPr${sorterAttrs.join(" ")}>`);
    parts.push(commonViewPrXml(opts.sorterView.view));
    parts.push("</p:sorterViewPr>");
  }

  // notesViewPr — sequence: cSldViewPr
  if (opts?.notesView) {
    parts.push("<p:notesViewPr>");
    parts.push(buildCSldViewPrXml());
    parts.push("</p:notesViewPr>");
  }

  parts.push(`<p:gridSpacing cx="${gridCx}" cy="${gridCy}"/>`);
  parts.push("</p:viewPr>");
  return parts.join("");
}
