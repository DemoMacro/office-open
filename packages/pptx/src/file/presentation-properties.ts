import type { Context } from "@file/xml-components";
import { ImportedXmlComponent } from "@file/xml-components";

import type { ShowOptions } from "./file";

function buildPresPropsXml(showOptions?: ShowOptions): string {
  const ns =
    'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
    'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
    'xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"';

  if (!showOptions) {
    return `<p:presentationPr ${ns}/>`;
  }

  // Build <p:showPr> element
  const showPrAttrs: string[] = [];
  if (showOptions.loop) showPrAttrs.push(' loop="1"');
  if (showOptions.showNarration === false) showPrAttrs.push(' showNarration="0"');
  if (showOptions.showAnimation === false) showPrAttrs.push(' showAnimation="0"');
  if (showOptions.useTimings) showPrAttrs.push(' useTimings="1"');

  // EG_ShowType: present | browse | kiosk
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

  // EG_SlideListChoice: sldAll | sldRg | custShow
  let slideListXml = "<p:sldAll/>";
  if (showOptions.slideRange) {
    slideListXml = `<p:sldRg st="${showOptions.slideRange.start}" end="${showOptions.slideRange.end}"/>`;
  }

  // penClr
  let penClrXml = "";
  if (showOptions.penColor) {
    penClrXml = `<p:penClr><a:srgbClr val="${showOptions.penColor}"/></p:penClr>`;
  }

  const showPrXml = `<p:showPr${showPrAttrs.join("")}>${showTypeXml}${slideListXml}${penClrXml}</p:showPr>`;

  return `<p:presentationPr ${ns}>${showPrXml}</p:presentationPr>`;
}

export class PresentationProperties extends ImportedXmlComponent {
  private static cache = new Map<string, ImportedXmlComponent>();
  private readonly key: string;

  public constructor(showOptions?: ShowOptions) {
    super("p:presentationPr");
    this.key = showOptions ? presPropsKey(showOptions) : "";
    if (!PresentationProperties.cache.has(this.key)) {
      PresentationProperties.cache.set(
        this.key,
        ImportedXmlComponent.fromXmlString(buildPresPropsXml(showOptions)),
      );
    }
  }

  public override toXml(context: Context): string {
    return PresentationProperties.cache.get(this.key)!.toXml(context);
  }
}

function presPropsKey(o: ShowOptions): string {
  return JSON.stringify(o);
}
