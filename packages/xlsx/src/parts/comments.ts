/**
 * Comments + VML notes descriptor for XLSX.
 *
 * Generates both xl/comments{n}.xml and xl/drawings/vmlDrawing{n}.vml
 * from the same CommentOptions array. Follows PPTX CustomDescriptor pattern.
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import { convertToPt } from "@office-open/core";
import { parseVmlShape } from "@office-open/core";
import { stringifyVmlShape } from "@office-open/core";
import { stringifyVmlShapetype } from "@office-open/core";
import { stringifyVmlShapeLayout } from "@office-open/core";
import type { LengthUnit, UniversalMeasure, VmlShapeStyle } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import type { WriteContext } from "@office-open/core/descriptor";
import { findChild, attr, textOf } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { escapeXml } from "@office-open/xml";

import { letterToColumn } from "../util/index";
import type {
  AnchorMarkerOptions,
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
      // commentPr is parsed but never re-emitted: Excel refuses to open a
      // third-party file carrying commentPr beside the VML note drawing this
      // compiler always writes (rival property systems — it reads the VML
      // shape's x:ClientData instead).
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

/** Per-note placement facts read from a vmlDrawing part (one per v:shape). */
export interface VmlNoteAnchor {
  /** 0-based row (x:Row). */
  row: number;
  /** 0-based column (x:Column). */
  column: number;
  /** x:Anchor 8-tuple when present. */
  anchor?: number[];
  /** Whether the note shape is visible (style visibility ≠ hidden). */
  visible: boolean;
  /** Shape width in points when present in the style. */
  width?: number;
  /** Shape height in points when present in the style. */
  height?: number;
}

export const vmlNotesDesc: CustomDescriptor<CommentsDocOptions, WriteContext, VmlNoteAnchor[]> = {
  kind: "custom",

  stringify(opts, _ctx) {
    if (opts.comments.length === 0) return undefined;

    const p: string[] = [
      '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">',
      stringifyVmlShapeLayout({ ext: "edit", idmap: { ext: "edit", data: "1" } }),
      stringifyVmlShapetype({
        id: "_x0000_t202",
        coordsize: "21600,21600",
        spt: 202,
        path: "m,l,21600r21600,l21600,xe",
        stroke: { joinstyle: "miter" },
        pathElement: { gradientshapeok: true, connecttype: "rect" },
      }),
    ];

    for (const [i, c] of opts.comments.entries()) {
      const { col, row } = cellRefToVmlCoords(c.cell);
      const anchor = c.anchor ?? [col, 0, row, 0, col + 2, 0, row + 2, 0];
      const style = {
        position: "absolute",
        marginLeft: "59.25pt",
        marginTop: "1.5pt",
        width: `${c.size?.width ?? DEFAULT_NOTE_WIDTH}pt` as UniversalMeasure,
        height: `${c.size?.height ?? DEFAULT_NOTE_HEIGHT}pt` as UniversalMeasure,
        zIndex: 1,
      } as VmlShapeStyle;
      if (!c.visible) style.visibility = "hidden";
      p.push(
        stringifyVmlShape({
          id: `_x0000_s${1025 + i}`,
          type: "#_x0000_t202",
          style,
          fillcolor: "infoBackground [80]",
          strokecolor: "none [81]",
          insetmode: "auto",
          fill: { color2: "infoBackground [80]" },
          shadow: { color: "none [81]", obscured: true },
          pathElement: { connecttype: "none" },
          textbox: {
            style: { directionAlt: "auto" },
            content: '<div style="text-align:left"></div>',
          },
          clientData: {
            objectType: "Note",
            MoveWithCells: "",
            SizeWithCells: "",
            Anchor: anchor.join(", "),
            AutoFill: false,
            Row: row,
            Column: col,
          },
        }),
      );
    }

    p.push("</xml>");
    return p.join("");
  },

  parse(el, _ctx) {
    const anchors: VmlNoteAnchor[] = [];
    for (const child of el.elements ?? []) {
      if (child.type !== "element" || child.name !== "v:shape") continue;
      const shape = parseVmlShape(child);
      const cd = shape.clientData;
      if (!cd || cd.objectType !== "Note" || cd.Row === undefined || cd.Column === undefined) {
        continue;
      }
      const note: VmlNoteAnchor = {
        row: cd.Row,
        column: cd.Column,
        visible: shape.style?.visibility !== "hidden",
      };
      const nums = (cd.Anchor ?? "")
        .split(/[,\s]+/)
        .filter(Boolean)
        .map(Number);
      if (nums.length === 8 && nums.every((n) => !Number.isNaN(n))) note.anchor = nums;
      const width = lengthToPt(shape.style?.width);
      if (width !== undefined) note.width = width;
      const height = lengthToPt(shape.style?.height);
      if (height !== undefined) note.height = height;
      anchors.push(note);
    }
    return anchors;
  },
};

/** Coerce a style length (number or measure string) to points; non-measure tokens yield undefined. */
function lengthToPt(value: LengthUnit | undefined): number | undefined {
  if (typeof value === "number") return value;
  if (typeof value !== "string" || !/^-?[\d.]/.test(value)) return undefined;
  return convertToPt(value as UniversalMeasure);
}

/** Excel's default note shape size (points). */
const DEFAULT_NOTE_WIDTH = 108;
const DEFAULT_NOTE_HEIGHT = 59.25;

/** Merge parsed VML note placement into comments by cell (column, row) pairing. */
export function mergeNoteAnchors(
  comments: CommentOptions[] | undefined,
  anchors: VmlNoteAnchor[],
): void {
  if (!comments || anchors.length === 0) return;
  for (const note of anchors) {
    const comment = comments.find((c) => {
      const { col, row } = cellRefToVmlCoords(c.cell);
      return col === note.column && row === note.row;
    });
    if (!comment) continue;
    if (note.anchor !== undefined) comment.anchor = note.anchor;
    if (note.visible) comment.visible = true;
    if (note.width !== undefined || note.height !== undefined) {
      comment.size = {
        width: note.width ?? DEFAULT_NOTE_WIDTH,
        height: note.height ?? DEFAULT_NOTE_HEIGHT,
      };
    }
  }
}

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

// ── Comment properties (CT_CommentPr) parse — stringify never emits it (see above) ──

function parseCommentPr(el: XmlElement): CommentPropertiesOptions {
  const pr: CommentPropertiesOptions = {};
  const locked = attr(el, "locked");
  if (locked !== undefined) pr.locked = parseOnOff(locked) ?? false;
  const defaultSize = attr(el, "defaultSize");
  if (defaultSize !== undefined) pr.defaultSize = parseOnOff(defaultSize) ?? false;
  const print = attr(el, "print");
  if (print !== undefined) pr.print = parseOnOff(print) ?? false;
  const disabled = attr(el, "disabled");
  if (disabled !== undefined) pr.disabled = parseOnOff(disabled) ?? false;
  const autoFill = attr(el, "autoFill");
  if (autoFill !== undefined) pr.autoFill = parseOnOff(autoFill) ?? false;
  const autoLine = attr(el, "autoLine");
  if (autoLine !== undefined) pr.autoLine = parseOnOff(autoLine) ?? false;
  const altText = attr(el, "altText");
  if (altText !== undefined) pr.altText = altText;
  const textHAlign = attr(el, "textHAlign");
  if (textHAlign !== undefined)
    pr.textHAlign = textHAlign as CommentPropertiesOptions["textHAlign"];
  const textVAlign = attr(el, "textVAlign");
  if (textVAlign !== undefined)
    pr.textVAlign = textVAlign as CommentPropertiesOptions["textVAlign"];
  const lockText = attr(el, "lockText");
  if (lockText !== undefined) pr.lockText = parseOnOff(lockText) ?? false;
  const justLastX = attr(el, "justLastX");
  if (justLastX !== undefined) pr.justLastX = parseOnOff(justLastX) ?? false;
  const autoScale = attr(el, "autoScale");
  if (autoScale !== undefined) pr.autoScale = parseOnOff(autoScale) ?? false;
  const anchorEl = findChild(el, "anchor");
  if (anchorEl) pr.anchor = parseAnchor(anchorEl);
  return pr;
}

function parseAnchor(el: XmlElement): ObjectAnchorOptions {
  const anchor: ObjectAnchorOptions = {
    moveWithCells: parseOnOff(attr(el, "moveWithCells")),
    sizeWithCells: parseOnOff(attr(el, "sizeWithCells")),
  };
  const from = findChild(el, "xdr:from");
  if (from) anchor.from = parseMarker(from);
  const to = findChild(el, "xdr:to");
  if (to) anchor.to = parseMarker(to);
  return anchor;
}

function parseMarker(el: XmlElement): AnchorMarkerOptions {
  const num = (tag: string) => Number(textOf(findChild(el, tag)!) ?? 0);
  return {
    col: num("xdr:col"),
    colOff: num("xdr:colOff"),
    row: num("xdr:row"),
    rowOff: num("xdr:rowOff"),
  };
}

// VML anchors use 0-based column/row; cell refs are 1-based uppercase letters + digits.
function cellRefToVmlCoords(ref: string): { col: number; row: number } {
  let i = 0;
  while (i < ref.length && ref.charCodeAt(i) >= 65 && ref.charCodeAt(i) <= 90) i++;
  return { col: letterToColumn(ref.slice(0, i)) - 1, row: parseInt(ref.slice(i), 10) - 1 };
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
