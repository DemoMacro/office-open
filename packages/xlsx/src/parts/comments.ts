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

import type {
  CommentOptions,
  CommentPropertiesOptions,
  ObjectAnchorOptions,
  RichTextOptions,
  RichTextRunOptions,
  RichTextRunPropertiesOptions,
} from "./worksheet";

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
      `<comments xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:xdr="http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing">`,
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
      // CT_Comment model is text then commentPr (sml.xsd:290-294).
      const commentPrXml = entry.commentPr ? buildCommentPrXml(entry.commentPr) : "";
      p.push(
        `<comment ref="${entry.cell}" authorId="${authorId}"><text>${textXml}</text>${commentPrXml}</comment>`,
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
        const commentPrEl = findChild(c, "commentPr");
        const comment: CommentOptions = {
          cell: ref,
          author: authors[authorId] ?? "",
          text,
        };
        if (commentPrEl) comment.commentPr = parseCommentPr(commentPrEl);
        comments.push(comment);
      }
    }

    return { comments } as CommentsDocOptions;
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
      const { col, row } = cellRefToVmlCoords(c.cell);
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
    return { comments: [] } as CommentsDocOptions;
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

// ── Comment properties (CT_CommentPr) helpers ──

/** Serialize CT_CommentPr. anchor is required (CT_ObjectAnchor), defaults to empty. */
function buildCommentPrXml(pr: CommentPropertiesOptions): string {
  const attrs: string[] = [];
  if (pr.locked !== undefined) attrs.push(`locked="${pr.locked ? 1 : 0}"`);
  if (pr.defaultSize !== undefined) attrs.push(`defaultSize="${pr.defaultSize ? 1 : 0}"`);
  if (pr.print !== undefined) attrs.push(`print="${pr.print ? 1 : 0}"`);
  if (pr.disabled !== undefined) attrs.push(`disabled="${pr.disabled ? 1 : 0}"`);
  if (pr.autoFill !== undefined) attrs.push(`autoFill="${pr.autoFill ? 1 : 0}"`);
  if (pr.autoLine !== undefined) attrs.push(`autoLine="${pr.autoLine ? 1 : 0}"`);
  if (pr.altText !== undefined) attrs.push(`altText="${escapeXml(pr.altText)}"`);
  if (pr.textHAlign !== undefined) attrs.push(`textHAlign="${pr.textHAlign}"`);
  if (pr.textVAlign !== undefined) attrs.push(`textVAlign="${pr.textVAlign}"`);
  if (pr.lockText !== undefined) attrs.push(`lockText="${pr.lockText ? 1 : 0}"`);
  if (pr.justLastX !== undefined) attrs.push(`justLastX="${pr.justLastX ? 1 : 0}"`);
  if (pr.autoScale !== undefined) attrs.push(`autoScale="${pr.autoScale ? 1 : 0}"`);
  return `<commentPr ${attrs.join(" ")}>${buildAnchorXml(pr.anchor)}</commentPr>`;
}

/** Serialize CT_ObjectAnchor. from/to/ext are optional in the XSD; attributes required. */
function buildAnchorXml(anchor: ObjectAnchorOptions | undefined): string {
  const moveWithCells = anchor?.moveWithCells ? 1 : 0;
  const sizeWithCells = anchor?.sizeWithCells ? 1 : 0;
  return `<xdr:anchor moveWithCells="${moveWithCells}" sizeWithCells="${sizeWithCells}"/>`;
}

function parseCommentPr(el: XmlElement): CommentPropertiesOptions {
  const pr: CommentPropertiesOptions = {};
  const locked = attr(el, "locked");
  if (locked !== undefined) pr.locked = locked === "1";
  const defaultSize = attr(el, "defaultSize");
  if (defaultSize !== undefined) pr.defaultSize = defaultSize === "1";
  const print = attr(el, "print");
  if (print !== undefined) pr.print = print === "1";
  const disabled = attr(el, "disabled");
  if (disabled !== undefined) pr.disabled = disabled === "1";
  const autoFill = attr(el, "autoFill");
  if (autoFill !== undefined) pr.autoFill = autoFill === "1";
  const autoLine = attr(el, "autoLine");
  if (autoLine !== undefined) pr.autoLine = autoLine === "1";
  const altText = attr(el, "altText");
  if (altText !== undefined) pr.altText = altText;
  const textHAlign = attr(el, "textHAlign");
  if (textHAlign !== undefined)
    pr.textHAlign = textHAlign as CommentPropertiesOptions["textHAlign"];
  const textVAlign = attr(el, "textVAlign");
  if (textVAlign !== undefined)
    pr.textVAlign = textVAlign as CommentPropertiesOptions["textVAlign"];
  const lockText = attr(el, "lockText");
  if (lockText !== undefined) pr.lockText = lockText === "1";
  const justLastX = attr(el, "justLastX");
  if (justLastX !== undefined) pr.justLastX = justLastX === "1";
  const autoScale = attr(el, "autoScale");
  if (autoScale !== undefined) pr.autoScale = autoScale === "1";
  const anchorEl = findChild(el, "xdr:anchor");
  if (anchorEl) pr.anchor = parseAnchor(anchorEl);
  return pr;
}

function parseAnchor(el: XmlElement): ObjectAnchorOptions {
  return {
    moveWithCells: attr(el, "moveWithCells") === "1",
    sizeWithCells: attr(el, "sizeWithCells") === "1",
  };
}

// VML anchors use 0-based column/row; cell refs are 1-based uppercase letters + digits.
function cellRefToVmlCoords(ref: string): { col: number; row: number } {
  let i = 0;
  while (i < ref.length && ref.charCodeAt(i) >= 65 && ref.charCodeAt(i) <= 90) i++;
  let col = 0;
  for (let j = 0; j < i; j++) col = col * 26 + (ref.charCodeAt(j) - 64);
  return { col: col - 1, row: parseInt(ref.slice(i), 10) - 1 };
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

/** Parse rich text element into a plain string or rich runs. */
function parseRst(textEl: XmlElement): string | RichTextOptions {
  const runs: RichTextRunOptions[] = [];
  const parts: string[] = [];
  let hasRuns = false;
  for (const child of textEl.elements ?? []) {
    if (child.name === "t") {
      parts.push(textOf(child) ?? "");
    } else if (child.name === "r") {
      hasRuns = true;
      const t = findChild(child, "t");
      const run: RichTextRunOptions = { text: t ? (textOf(t) ?? "") : "" };
      const rPr = findChild(child, "rPr");
      if (rPr) {
        const props: RichTextRunPropertiesOptions = {};
        if (findChild(rPr, "b")) props.bold = true;
        if (findChild(rPr, "i")) props.italic = true;
        const uEl = findChild(rPr, "u");
        if (uEl)
          props.underline =
            (attr(uEl, "val") as RichTextRunPropertiesOptions["underline"]) ?? "single";
        if (findChild(rPr, "strike")) props.strike = true;
        const szEl = findChild(rPr, "sz");
        if (szEl) {
          const sz = Number(attr(szEl, "val"));
          if (!Number.isNaN(sz)) props.size = sz;
        }
        const colorEl = findChild(rPr, "color");
        if (colorEl && attr(colorEl, "rgb")) props.color = attr(colorEl, "rgb");
        const rFontEl = findChild(rPr, "rFont");
        if (rFontEl && attr(rFontEl, "val")) props.font = attr(rFontEl, "val");
        run.properties = props;
      }
      runs.push(run);
    }
  }
  if (hasRuns) return { runs };
  return parts.join("");
}
