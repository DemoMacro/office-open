import { BaseXmlComponent } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { escapeXml } from "@office-open/xml";

import { buildRstXml } from "./shared-strings";
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
      const textXml =
        typeof entry.text === "string"
          ? `<t>${escapeXml(entry.text)}</t>`
          : buildRstXml(entry.text);
      // commentPr (CT_CommentPr, optional)
      let commentPrXml = "";
      if (entry.commentPr) {
        const cp = entry.commentPr;
        const cpAttrs: string[] = [];
        if (cp.locked === false) cpAttrs.push('locked="0"');
        if (cp.defaultSize === false) cpAttrs.push('defaultSize="0"');
        if (cp.print === false) cpAttrs.push('print="0"');
        if (cp.disabled) cpAttrs.push('disabled="1"');
        if (cp.autoFill === false) cpAttrs.push('autoFill="0"');
        if (cp.autoLine === false) cpAttrs.push('autoLine="0"');
        if (cp.altText) cpAttrs.push(`altText="${escapeXml(cp.altText)}"`);
        if (cp.textHAlign && cp.textHAlign !== "left")
          cpAttrs.push(`textHAlign="${cp.textHAlign}"`);
        if (cp.textVAlign && cp.textVAlign !== "top") cpAttrs.push(`textVAlign="${cp.textVAlign}"`);
        if (cp.lockText === false) cpAttrs.push('lockText="0"');
        if (cp.justLastX) cpAttrs.push('justLastX="1"');
        if (cp.autoScale) cpAttrs.push('autoScale="1"');
        // anchor (CT_ObjectAnchor, required child of commentPr)
        const anchorAttrs: string[] = [];
        if (cp.anchor?.moveWithCells) anchorAttrs.push('moveWithCells="1"');
        if (cp.anchor?.sizeWithCells) anchorAttrs.push('sizeWithCells="1"');
        commentPrXml = `<commentPr${cpAttrs.length ? " " + cpAttrs.join(" ") : ""}><anchor${anchorAttrs.length ? " " + anchorAttrs.join(" ") : ""}/></commentPr>`;
      }
      p.push(
        `<comment ref="${entry.cell}" authorId="${authorId}"><text>${textXml}</text>${commentPrXml}</comment>`,
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
