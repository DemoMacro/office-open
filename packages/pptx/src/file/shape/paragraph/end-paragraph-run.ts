import { BuilderElement } from "@file/xml-components";
import type { Context } from "@file/xml-components";

/**
 * a:endParaRPr — End paragraph run properties.
 * Lazy: stores lang, builds XML in toXml().
 */
export class EndParagraphRunProperties extends BuilderElement<{
  readonly lang: string;
}> {
  private readonly lang: string;

  public constructor(lang: string = "en-US") {
    super({ name: "a:endParaRPr" });
    this.lang = lang;
  }

  public override toXml(_context: Context): string {
    return `<a:endParaRPr lang="${this.lang}"/>`;
  }
}
