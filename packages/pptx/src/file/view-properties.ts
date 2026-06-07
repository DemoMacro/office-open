type SplitterBarState = "restored" | "maximized" | "minimized";

export interface NormalViewOptions {
  showOutlineIcons?: boolean;
  snapVertSplitter?: boolean;
  vertBarState?: SplitterBarState;
  horzBarState?: SplitterBarState;
  preferSingleView?: boolean;
}

export interface SlideViewOptions {
  snapToGrid?: boolean;
  snapToObjects?: boolean;
  showGuides?: boolean;
  varScale?: boolean;
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
  zoomScaleNumerator?: number;
  zoomScaleDenominator?: number;
  normalView?: NormalViewOptions;
  slideView?: SlideViewOptions;
  guides?: {
    orient?: "vert" | "horz";
    pos?: number;
  }[];
  outlineView?: {
    slides?: {
      rId: string;
      collapse?: boolean;
    }[];
  };
  sorterView?: {
    showFormatting?: boolean;
  };
  notesView?: boolean;
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

function buildCViewPrXml(zoomN: number, zoomD: number, varScale = false): string {
  const cViewPrAttr = varScale ? ' varScale="1"' : "";
  return (
    `<p:cViewPr${cViewPrAttr}>` +
    `<p:scale><a:sx n="${zoomN}" d="${zoomD}"/><a:sy n="${zoomN}" d="${zoomD}"/></p:scale>` +
    '<p:origin x="1200" y="72"/>' +
    "</p:cViewPr>"
  );
}

function buildCSldViewPrXml(
  opts?: SlideViewOptions,
  zoomN = 90,
  zoomD = 100,
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
    buildCViewPrXml(zoomN, zoomD, opts?.varScale !== false) +
    guideLstXml +
    "</p:cSldViewPr>"
  );
}

export function buildViewPropsXml(opts?: ViewPropertiesOptions): string {
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
  parts.push(buildCSldViewPrXml(opts?.slideView, zoomN, zoomD, opts?.guides));
  parts.push("</p:slideViewPr>");

  // outlineViewPr (D) — sequence: cViewPr, sldLst
  if (opts?.outlineView) {
    let outlineXml = "<p:outlineViewPr>";
    outlineXml += buildCViewPrXml(zoomN, zoomD);
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

  parts.push("<p:notesTextViewPr>");
  parts.push("<p:cViewPr>");
  parts.push('<p:scale><a:sx n="1" d="1"/><a:sy n="1" d="1"/></p:scale>');
  parts.push('<p:origin x="0" y="0"/>');
  parts.push("</p:cViewPr>");
  parts.push("</p:notesTextViewPr>");

  // sorterViewPr (D) — sequence: cViewPr (not cSldViewPr!)
  if (opts?.sorterView) {
    const sorterAttrs: string[] = [];
    if (opts.sorterView.showFormatting) sorterAttrs.push(' showFormatting="1"');
    parts.push(`<p:sorterViewPr${sorterAttrs.join("")}>`);
    parts.push(buildCViewPrXml(zoomN, zoomD));
    parts.push("</p:sorterViewPr>");
  }

  // notesViewPr (D) — sequence: cSldViewPr
  if (opts?.notesView) {
    parts.push("<p:notesViewPr>");
    parts.push(buildCSldViewPrXml(undefined, zoomN, zoomD));
    parts.push("</p:notesViewPr>");
  }

  parts.push(`<p:gridSpacing cx="${gridCx}" cy="${gridCy}"/>`);
  parts.push("</p:viewPr>");
  return parts.join("");
}
