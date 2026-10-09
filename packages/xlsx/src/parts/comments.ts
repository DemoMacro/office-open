/**
 * Comments + VML notes descriptor for XLSX.
 *
 * Generates both xl/comments{n}.xml and xl/drawings/vmlDrawing{n}.vml
 * from the same CommentOptions array. Follows PPTX CustomDescriptor pattern.
 *
 * @module
 */

import { parseOnOff } from "@office-open/core";
import { convertToEmu, convertToPt } from "@office-open/core";
import { parseVmlShape } from "@office-open/core";
import { parseVmlShapeLayout } from "@office-open/core";
import { parseVmlShapetype } from "@office-open/core";
import { stringifyVmlShape } from "@office-open/core";
import { stringifyVmlShapetype } from "@office-open/core";
import { stringifyVmlShapeLayout } from "@office-open/core";
import type {
  LengthUnit,
  UniversalMeasure,
  VmlShapeLayoutOptions,
  VmlShapeOptions,
  VmlShapetypeOptions,
  VmlShapeStyle,
} from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import type { WriteContext } from "@office-open/core/descriptor";
import { attr, attrNum, findChild, textOf } from "@office-open/xml";
import type { Element as XmlElement } from "@office-open/xml";
import { escapeXml } from "@office-open/xml";

import { parseA1Cell } from "../util/index";
import { buildRstXml, parseRPr, tElement } from "./shared-strings";
import type {
  AnchorMarkerOptions,
  CommentOptions,
  CommentPropertiesOptions,
  NoteAnchorOptions,
  ObjectAnchorOptions,
  RichTextOptions,
  RichTextRunOptions,
} from "./worksheet";

// ── Comments descriptor (xl/comments{n}.xml) ──

export interface CommentsDocOptions {
  comments: CommentOptions[];
}

export interface CommentsVmlOptions extends CommentsDocOptions {
  /** Source o:shapelayout settings; fresh comments use Excel's idmap default. */
  layout?: VmlShapeLayoutOptions;
  /** Source v:shapetype; fresh comments use the standard note-callout shape. */
  shapeType?: VmlShapetypeOptions;
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
        typeof entry.text === "string" ? tElement(entry.text) : buildRstXml(entry.text);
      const propertiesXml = entry.properties ? buildCommentPrXml(entry.properties) : "";
      const shapeId = entry.shapeId !== undefined ? ` shapeId="${entry.shapeId}"` : "";
      const uidAttr = entry.uid !== undefined ? ` xr:uid="${escapeXml(entry.uid)}"` : "";
      p.push(
        `<comment ref="${entry.cell}"${shapeId}${uidAttr} authorId="${authorId}"><text>${textXml}</text>${propertiesXml}</comment>`,
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
        const shapeId = attrNum(c, "shapeId");
        const uid = attr(c, "xr:uid");
        const authorId = Number(attr(c, "authorId") ?? 0);
        const textEl = findChild(c, "text");
        const text = textEl ? parseRst(textEl) : "";
        const commentPrEl = findCommentPrElement(c);
        const comment: CommentOptions = {
          cell: ref,
          ...(shapeId !== undefined ? { shapeId } : {}),
          ...(uid !== undefined ? { uid } : {}),
          author: authors[authorId] ?? "",
          text,
        };
        if (commentPrEl) comment.properties = parseCommentPr(commentPrEl);
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
  /** Cell-corner anchor (x:Anchor) when present. */
  anchor?: NoteAnchorOptions;
  /** Whether the note shape is visible (style visibility ≠ hidden). */
  visible: boolean;
  /** Shape width in points when present in the style. */
  width?: number;
  /** Shape height in points when present in the style. */
  height?: number;
  /** Full source note shape, including unprojected VML fidelity. */
  shape: VmlShapeOptions;
}

export interface VmlNotesParseResult {
  layout?: VmlShapeLayoutOptions;
  shapeType?: VmlShapetypeOptions;
  anchors: VmlNoteAnchor[];
}

export const vmlNotesDesc: CustomDescriptor<CommentsVmlOptions, WriteContext, VmlNotesParseResult> =
  {
    kind: "custom",

    stringify(opts, _ctx) {
      if (opts.comments.length === 0) return undefined;

      const p: string[] = [
        '<xml xmlns:v="urn:schemas-microsoft-com:vml" xmlns:o="urn:schemas-microsoft-com:office:office" xmlns:x="urn:schemas-microsoft-com:office:excel">',
        stringifyVmlShapeLayout(opts.layout ?? { ext: "edit", idmap: { ext: "edit", data: "1" } }),
        stringifyVmlShapetype(
          opts.shapeType ?? {
            id: "_x0000_t202",
            coordsize: "21600,21600",
            spt: 202,
            path: "m,l,21600r21600,l21600,xe",
            stroke: { joinstyle: "miter" },
            pathElement: { gradientshapeok: true, connecttype: "rect" },
          },
        ),
      ];

      for (const [i, c] of opts.comments.entries()) {
        const { col, row } = cellRefToVmlCoords(c.cell);
        const shape = c.vmlShape ? { ...c.vmlShape } : defaultVmlNoteShape(c, col, row, i);
        if (c.vmlShape) applyNoteProjections(shape, c, col, row);
        p.push(stringifyVmlShape(shape));
      }

      p.push("</xml>");
      return p.join("");
    },

    parse(el, _ctx) {
      const anchors: VmlNoteAnchor[] = [];
      let layout: VmlShapeLayoutOptions | undefined;
      let shapeType: VmlShapetypeOptions | undefined;
      for (const child of el.elements ?? []) {
        if (child.type === "element" && child.name === "o:shapelayout") {
          layout = parseVmlShapeLayout(child);
          continue;
        }
        if (child.type === "element" && child.name === "v:shapetype") {
          shapeType = parseVmlShapetype(child);
          continue;
        }
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
          shape,
        };
        const nums = (cd.Anchor ?? "")
          .split(/[,\s]+/)
          .filter(Boolean)
          .map(Number);
        if (nums.length === 8 && nums.every((n) => !Number.isNaN(n))) {
          // VML stores pixel offsets; the public API carries EMU (px × 9525).
          note.anchor = {
            from: {
              col: nums[0]!,
              colOff: nums[1]! * 9525,
              row: nums[2]!,
              rowOff: nums[3]! * 9525,
            },
            to: { col: nums[4]!, colOff: nums[5]! * 9525, row: nums[6]!, rowOff: nums[7]! * 9525 },
          };
        }
        const width = lengthToPt(shape.style?.width);
        if (width !== undefined) note.width = width;
        const height = lengthToPt(shape.style?.height);
        if (height !== undefined) note.height = height;
        anchors.push(note);
      }
      return { anchors, ...(layout ? { layout } : {}), ...(shapeType ? { shapeType } : {}) };
    },
  };

function defaultVmlNoteShape(
  comment: CommentOptions,
  col: number,
  row: number,
  index: number,
): VmlShapeOptions {
  const style = {
    position: "absolute",
    marginLeft: "59.25pt",
    marginTop: "1.5pt",
    width: `${comment.size?.width ?? DEFAULT_NOTE_WIDTH}pt` as UniversalMeasure,
    height: `${comment.size?.height ?? DEFAULT_NOTE_HEIGHT}pt` as UniversalMeasure,
    zIndex: 1,
  } as VmlShapeStyle;
  if (!comment.visible) style.visibility = "hidden";
  return {
    id: `_x0000_s${1025 + index}`,
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
      Anchor: noteAnchorValues(comment, col, row).join(", "),
      AutoFill: false,
      Row: row,
      Column: col,
    },
  };
}

function applyNoteProjections(
  shape: VmlShapeOptions,
  comment: CommentOptions,
  col: number,
  row: number,
): void {
  const style = { ...shape.style };
  if (comment.size) {
    style.width = `${comment.size.width}pt`;
    style.height = `${comment.size.height}pt`;
  }
  if (comment.visible) delete style.visibility;
  else style.visibility = "hidden";
  shape.style = style;
  const clientData = shape.clientData ?? { objectType: "Note" as const };
  shape.clientData = {
    ...clientData,
    ...(comment.anchor ? { Anchor: noteAnchorValues(comment, col, row).join(", ") } : {}),
    Row: row,
    Column: col,
  };
}

function noteAnchorValues(comment: CommentOptions, col: number, row: number): number[] {
  if (!comment.anchor) return [col, 0, row, 0, col + 2, 0, row + 2, 0];
  return [
    comment.anchor.from.col,
    Math.round(convertToEmu(comment.anchor.from.colOff ?? 0) / 9525),
    comment.anchor.from.row,
    Math.round(convertToEmu(comment.anchor.from.rowOff ?? 0) / 9525),
    comment.anchor.to.col,
    Math.round(convertToEmu(comment.anchor.to.colOff ?? 0) / 9525),
    comment.anchor.to.row,
    Math.round(convertToEmu(comment.anchor.to.rowOff ?? 0) / 9525),
  ];
}

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
    if (note.shape) comment.vmlShape = note.shape;
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

// ── Comment properties (CT_CommentPr) ──

/**
 * Modern producers may wrap CT_CommentPr in mc:AlternateContent. The
 * transitional comments schema models the direct/fallback element; unknown
 * Choice extensions are used only when no canonical property element exists.
 */
function findCommentPrElement(comment: XmlElement): XmlElement | undefined {
  return (
    findChild(comment, "commentPr") ??
    (comment.elements ?? [])
      .flatMap((alternate) => {
        if (alternate.name !== "mc:AlternateContent") return [];
        const fallback = findChild(alternate, "mc:Fallback");
        const choice = findChild(alternate, "mc:Choice");
        return [
          fallback ? findChild(fallback, "commentPr") : undefined,
          choice ? findChild(choice, "commentPr") : undefined,
        ];
      })
      .find((element) => element !== undefined)
  );
}

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
  const colHidden = attr(el, "colHidden");
  if (colHidden !== undefined) pr.colHidden = parseOnOff(colHidden) ?? false;
  const rowHidden = attr(el, "rowHidden");
  if (rowHidden !== undefined) pr.rowHidden = parseOnOff(rowHidden) ?? false;
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
  if (anchorEl) {
    const anchor = parseAnchor(anchorEl);
    if (anchor) pr.anchor = anchor;
  }
  return pr;
}

/** Serialize CT_CommentPr with XSD attribute order and required anchor. */
function buildCommentPrXml(pr: CommentPropertiesOptions): string {
  const attrs: string[] = [];
  const bool = (name: string, value: boolean | undefined): void => {
    if (value !== undefined) attrs.push(` ${name}="${value ? 1 : 0}"`);
  };
  bool("locked", pr.locked);
  bool("defaultSize", pr.defaultSize);
  bool("print", pr.print);
  bool("disabled", pr.disabled);
  bool("autoFill", pr.autoFill);
  bool("autoLine", pr.autoLine);
  bool("colHidden", pr.colHidden);
  bool("rowHidden", pr.rowHidden);
  if (pr.altText !== undefined) attrs.push(` altText="${escapeXml(pr.altText)}"`);
  if (pr.textHAlign !== undefined) attrs.push(` textHAlign="${pr.textHAlign}"`);
  if (pr.textVAlign !== undefined) attrs.push(` textVAlign="${pr.textVAlign}"`);
  bool("lockText", pr.lockText);
  bool("justLastX", pr.justLastX);
  bool("autoScale", pr.autoScale);
  const anchor = pr.anchor;
  const anchorXml = anchor
    ? `<anchor${boolAttr("moveWithCells", anchor.moveWithCells)}${boolAttr(
        "sizeWithCells",
        anchor.sizeWithCells,
      )}>${markerXml("xdr:from", anchor.from)}${markerXml("xdr:to", anchor.to)}</anchor>`
    : "";
  return `<commentPr${attrs.join("")}>${anchorXml}</commentPr>`;
}

function boolAttr(name: string, value: boolean | undefined): string {
  return value !== undefined ? ` ${name}="${value ? 1 : 0}"` : "";
}

function markerXml(name: string, marker: AnchorMarkerOptions): string {
  const offset = (tag: string, value: number | undefined): string =>
    value !== undefined ? `<${tag}>${value}</${tag}>` : "";
  return (
    `<${name}><xdr:col>${marker.col}</xdr:col>${offset("xdr:colOff", marker.colOff)}` +
    `<xdr:row>${marker.row}</xdr:row>${offset("xdr:rowOff", marker.rowOff)}</${name}>`
  );
}

// CT_ObjectAnchor's from/to are required in the XSD; a malformed element
// missing either corner yields undefined.
function parseAnchor(el: XmlElement): ObjectAnchorOptions | undefined {
  const from = findChild(el, "xdr:from") ?? findChild(el, "from");
  const to = findChild(el, "xdr:to") ?? findChild(el, "to");
  if (!from || !to) return undefined;
  return {
    moveWithCells: parseOnOff(attr(el, "moveWithCells")),
    sizeWithCells: parseOnOff(attr(el, "sizeWithCells")),
    from: parseMarker(from),
    to: parseMarker(to),
  };
}

function parseMarker(el: XmlElement): AnchorMarkerOptions {
  const num = (tag: string) => Number(textOf(findChild(el, tag)!) ?? 0);
  return {
    col: num("xdr:col") ?? 0,
    colOff: num("xdr:colOff"),
    row: num("xdr:row") ?? 0,
    rowOff: num("xdr:rowOff"),
  };
}

// VML anchors use 0-based column/row; cell refs are 1-based uppercase letters + digits.
function cellRefToVmlCoords(ref: string): { col: number; row: number } {
  const cell = parseA1Cell(ref);
  return { col: (cell?.col ?? 1) - 1, row: (cell?.row ?? 1) - 1 };
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
      if (t?.attributes?.["xml:space"] !== undefined) {
        run.textSpaceRaw = String(t.attributes["xml:space"]);
      }
      const rPr = findChild(child, "rPr");
      if (rPr) run.properties = parseRPr(rPr);
      runs.push(run);
    }
  }
  if (hasRuns) return { runs };
  return parts.join("");
}
