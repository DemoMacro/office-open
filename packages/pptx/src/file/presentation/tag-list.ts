import type { Context } from "@file/xml-components";
import { BaseXmlComponent } from "@file/xml-components";

export interface TagOptions {
  readonly name: string;
  readonly val: string;
}

/**
 * p:tagLst — Tag list part content (CT_TagList).
 * Contains string tags referenced via p:custDataLst > p:tags in the presentation.
 */
export class TagList extends BaseXmlComponent {
  private readonly tags: readonly TagOptions[];

  public constructor(tags: readonly TagOptions[]) {
    super("p:tagLst");
    this.tags = tags;
  }

  public override toXml(_context: Context): string {
    if (this.tags.length === 0) return "<p:tagLst/>";
    const tagXmls = this.tags.map((t) => `<p:tag name="${t.name}" val="${t.val}"/>`).join("");
    return `<p:tagLst>${tagXmls}</p:tagLst>`;
  }
}
