/**
 * Comments + VML notes descriptor for XLSX.
 *
 * Generates both xl/comments{n}.xml and xl/drawings/vmlDrawing{n}.vml
 * from the same CommentOptions array. Follows PPTX CustomDescriptor pattern.
 *
 * @module
 */

import type { CustomDescriptor } from "@office-open/core/descriptor";
import { findChild, attr, textOf } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { escapeXml } from "@office-open/xml";

import type { CommentOptions, RichTextOptions } from "./worksheet";

// ── Comments descriptor (xl/comments{n}.xml) ──

export interface CommentsDocOptions {
  comments: CommentOptions[];
}

export const commentsDesc: CustomDescriptor<CommentsDocOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    if (opts.comments.length === 0) return undefined;
    const authors = collectAuthors(opts.comments);
    const p: string[] = [
      `<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">`,
      `<authors>`,
    ];

    for (const author of authors) {
      p.push(`<author>${escapeXml(author)}</author>`);
    }

    p.push("</authors><commentList>");

    for (const entry of opts.comments) {
      const authorId = authors.indexOf(entry.author);
      const textXml =
        typeof entry.text === "string"
          ? `<t>${escapeXml(entry.text)}</t>`
          : buildRstXml(entry.text);
      p.push(
        `<comment ref="${entry.cell}" authorId="${authorId}"><text>${textXml}</text></comment>`,
      );
    }

    p.push("</commentList></comments>");
    return p.join("");
  },

  parse(el, _ctx) {
    const comments: CommentOptions[] = [];
    const authors: string[] = [];

    const authorsEl = findChild(el, "authors");
    if (authorsEl) {
      for (const a of authorsEl.elements ?? []) {
        if (a.name === "author") authors.push(textOf(a) ?? "");
      }
    }

    const listEl = findChild(el, "commentList");
    if (listEl) {
      for (const c of listEl.elements ?? []) {
        if (c.name !== "comment") continue;
        const ref = attr(c, "ref") ?? "";
        const authorId = Number(attr(c, "authorId") ?? 0);
        const textEl = findChild(c, "text");
        const text = textEl ? parseRst(textEl) : "";
        comments.push({
          cell: ref,
          author: authors[authorId] ?? "",
          text,
        });
      }
    }

    return { comments } as Record<string, unknown>;
  },
};

// ── VML notes descriptor (xl/drawings/vmlDrawing{n}.vml) ──

export const vmlNotesDesc: CustomDescriptor<CommentsDocOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    if (opts.comments.length === 0) return undefined;

    const p: string[] = [
      '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">',
      '<o:shapelayout v:ext="edit"><o:idmap v:ext="edit" data="1"/></o:shapelayout>',
      '<v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202" path="m,l,21600r21600,l21600,xe">',
      '<v:stroke joinstyle="miter"/>',
      '<v:path gradientshapeok="t" o:connecttype="rect"/>',
      "</v:shapetype>",
    ];

    for (let i = 0; i < opts.comments.length; i++) {
      const c = opts.comments[i];
      const col = c.cell.charCodeAt(0) - 65;
      const row = parseInt(c.cell.slice(1), 10) - 1;
      const anchor = `${col}, 0, ${row}, 0, ${col + 2}, 0, ${row + 2}, 0`;
      p.push(
        `<v:shape id="_x0000_s${1025 + i}" type="#_x0000_t202" ` +
          `style="position:absolute;margin-left:59.25pt;margin-top:1.5pt;width:108pt;height:59.25pt;` +
          `z-index:1;visibility:hidden" fillcolor="infoBackground [80]" strokecolor="none [81]" o:insetmode="auto">`,
        `<v:fill color2="infoBackground [80]"/>`,
        `<v:shadow color="none [81]" obscured="t"/>`,
        `<v:path o:connecttype="none"/>`,
        `<v:textbox style="mso-direction-alt:auto"><div style="text-align:left"></div></v:textbox>`,
        `<x:ClientData ObjectType="Note"><x:MoveWithCells/><x:SizeWithCells/>`,
        `<x:Anchor>${anchor}</x:Anchor>`,
        `<x:AutoFill>False</x:AutoFill>`,
        `<x:Row>${row}</x:Row>`,
        `<x:Column>${col}</x:Column>`,
        `</x:ClientData>`,
        `</v:shape>`,
      );
    }

    p.push("</xml>");
    return p.join("");
  },

  parse(_el, _ctx) {
    // VML parsing is not commonly needed — return empty
    return { comments: [] } as Record<string, unknown>;
  },
};

// ── Helpers ──

function collectAuthors(comments: CommentOptions[]): string[] {
  const seen = new Set<string>();
  const result: string[] = [];
  for (const entry of comments) {
    if (!seen.has(entry.author)) {
      seen.add(entry.author);
      result.push(entry.author);
    }
  }
  return result.length > 0 ? result : [""];
}

/** Build rich text (CT_Rst) XML from runs. */
function buildRstXml(rst: RichTextOptions): string {
  const runs = rst.runs ?? [];
  const parts: string[] = [];
  for (const run of runs) {
    const props = run.properties;
    if (!props) {
      parts.push(`<r><t>${escapeXml(run.text)}</t></r>`);
      continue;
    }
    const rPr: string[] = [];
    if (props.bold) rPr.push("<b/>");
    if (props.italic) rPr.push("<i/>");
    if (props.underline) rPr.push(`<u val="${props.underline}"/>`);
    if (props.strike) rPr.push("<strike/>");
    if (props.size) rPr.push(`<sz val="${props.size}"/>`);
    if (props.color) rPr.push(`<color rgb="${props.color}"/>`);
    if (props.font) rPr.push(`<rFont val="${props.font}"/>`);
    const rPrXml = rPr.length ? `<rPr>${rPr.join("")}</rPr>` : "";
    parts.push(`<r>${rPrXml}<t>${escapeXml(run.text)}</t></r>`);
  }
  return parts.join("");
}

/** Parse rich text element into a simple string. */
function parseRst(textEl: XmlElement): string {
  const parts: string[] = [];
  for (const child of textEl.elements ?? []) {
    if (child.name === "t") {
      parts.push(textOf(child) ?? "");
    } else if (child.name === "r") {
      const t = findChild(child, "t");
      if (t) parts.push(textOf(t) ?? "");
    }
  }
  return parts.join("");
}
