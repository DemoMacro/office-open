import type { Context } from "@file/xml-components";
import { ImportedXmlComponent } from "@file/xml-components";

export interface ShowOptions {
  readonly loop?: boolean;
  readonly kiosk?: boolean;
  readonly showNarration?: boolean;
  readonly useTimings?: boolean;
}

function buildPresPropsXml(showOptions?: ShowOptions): string {
  if (!showOptions) {
    return `<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main"/>`;
  }

  const attrs: string[] = [];
  if (showOptions.loop) attrs.push(' loop="1"');
  if (showOptions.showNarration === false) attrs.push(' showNarration="0"');
  if (showOptions.useTimings) attrs.push(' useTimings="1"');

  // EG_ShowType: present | browse | kiosk (child element, not attribute)
  const showType = showOptions.kiosk ? "<p:kiosk/>" : "<p:present/>";
  const showPrXml = `<p:showPr${attrs.join("")}>${showType}</p:showPr>`;

  return `<p:presentationPr xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:p="http://schemas.openxmlformats.org/presentationml/2006/main">${showPrXml}</p:presentationPr>`;
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
  return `l${o.loop ? 1 : 0}k${o.kiosk ? 1 : 0}n${o.showNarration ? 1 : 0}t${o.useTimings ? 1 : 0}`;
}
