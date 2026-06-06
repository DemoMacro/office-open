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
    const hasCommentPr = this.entries.some((e) => e.commentPr);
    const nsAttrs = hasCommentPr
      ? ' xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"'
      : "";
    const p: string[] = [
      `<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main"${nsAttrs}>`,
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
      const commentPrXml = this.buildCommentPrXml(entry);
      p.push(
        `<comment ref="${entry.cell}" authorId="${authorId}"><text>${textXml}</text>${commentPrXml}</comment>`,
      );
    }

    p.push("</commentList></comments>");
    return p.join("");
  }

  /** Build CT_CommentPr XML. Returns empty string if commentPr is not set. */
  private buildCommentPrXml(entry: CommentOptions): string {
    const cp = entry.commentPr;
    if (!cp) return "";

    const attrs: string[] = [];
    if (cp.locked === false) attrs.push('locked="0"');
    if (cp.defaultSize === false) attrs.push('defaultSize="0"');
    if (cp.print === false) attrs.push('print="0"');
    if (cp.disabled) attrs.push('disabled="1"');
    if (cp.autoFill === false) attrs.push('autoFill="0"');
    if (cp.autoLine === false) attrs.push('autoLine="0"');
    if (cp.altText) attrs.push(`altText="${escapeXml(cp.altText)}"`);
    if (cp.textHAlign && cp.textHAlign !== "left") attrs.push(`textHAlign="${cp.textHAlign}"`);
    if (cp.textVAlign && cp.textVAlign !== "top") attrs.push(`textVAlign="${cp.textVAlign}"`);
    if (cp.lockText === false) attrs.push('lockText="0"');
    if (cp.justLastX) attrs.push('justLastX="1"');
    if (cp.autoScale) attrs.push('autoScale="1"');

    // CT_ObjectAnchor requires xdr:from and xdr:to children.
    // Derive position from the cell reference (same as VML).
    const anchorXml = this.buildAnchorXml(entry.cell, cp.anchor);
    const attrStr = attrs.length ? ` ${attrs.join(" ")}` : "";
    return `<commentPr${attrStr}>${anchorXml}</commentPr>`;
  }

  /** Build CT_ObjectAnchor with required xdr:from/xdr:to markers. */
  private buildAnchorXml(
    cell: string,
    anchor: NonNullable<CommentOptions["commentPr"]>["anchor"],
  ): string {
    const col = cell.charCodeAt(0) - 65;
    const row = parseInt(cell.slice(1), 10) - 1;
    // Default anchor: from the comment cell to 2 cols/rows right/down
    const toCol = col + 2;
    const toRow = row + 2;

    const anchorAttrs: string[] = [];
    if (anchor?.moveWithCells) anchorAttrs.push(' moveWithCells="1"');
    if (anchor?.sizeWithCells) anchorAttrs.push(' sizeWithCells="1"');

    return (
      `<anchor${anchorAttrs.join("")}>` +
      `<xdr:from><xdr:col>${col}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${row}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>` +
      `<xdr:to><xdr:col>${toCol}</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>${toRow}</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>` +
      `</anchor>`
    );
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
