import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { escapeXml } from "@office-open/xml";

import type { CommentOptions } from "./worksheet";

/**
 * Generates xl/comments{n}.xml — cell comment data.
 *
 * @module
 */
export class Comments extends BaseXmlComponent {
  public constructor(private readonly entries: readonly CommentOptions[]) {
    super("comments");
  }

  public override toXml(_context: Context): string {
    const authors = this.collectAuthors();
    const p: string[] = [
      '<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">',
      `<authors>`,
    ];

    for (const author of authors) {
      p.push(`<author>${escapeXml(author)}</author>`);
    }

    p.push("</authors><commentList>");

    for (const entry of this.entries) {
      const authorId = authors.indexOf(entry.author);
      p.push(
        `<comment ref="${entry.cell}" authorId="${authorId}"><text><t>${escapeXml(entry.text)}</t></text></comment>`,
      );
    }

    p.push("</commentList></comments>");
    return p.join("");
  }

  private collectAuthors(): string[] {
    const seen = new Set<string>();
    const result: string[] = [];
    for (const entry of this.entries) {
      if (!seen.has(entry.author)) {
        seen.add(entry.author);
        result.push(entry.author);
      }
    }
    return result.length > 0 ? result : [""];
  }
}
