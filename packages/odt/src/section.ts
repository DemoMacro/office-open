import type { OfficeFormsOptions } from "@office-open/core";
import type { DocumentOptions, ParagraphChild, SectionChild } from "@office-open/docx";
import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  escapeText,
  textOf,
  xmlElement,
} from "@office-open/odf";
import type { Element } from "@office-open/xml";

import { blockXml, blocksXml, parseBlocks, unknownOdtElement, type ParseContext } from "./body";
import type { OdtChart, OdtImage } from "./drawing";
import type { OdtEmbeddedObjectOptions } from "./embedded-objects";
import { OdtParseError } from "./error";
import type { CanonicalRevision, OdtVariableValueType } from "./paragraph";

/** Footnote/endnote entry, indexed from the shared docx document model. */
export type NoteEntry = NonNullable<DocumentOptions["footnotes"]>[number];

export type NoteChildren = NoteEntry["children"];

export type FootnoteProperties = NonNullable<
  NonNullable<DocumentOptions["settings"]>["footnoteProperties"]
>;

export type EndnoteProperties = NonNullable<
  NonNullable<DocumentOptions["settings"]>["endnoteProperties"]
>;

/** docx ST_NumberFormat tokens with a direct ODF num-format token. */
export const NUM_FORMAT_ODF: Record<string, string> = {
  decimal: "1",
  lowerLetter: "a",
  upperLetter: "A",
  lowerRoman: "i",
  upperRoman: "I",
};

export const NUM_FORMAT_DOCX: Record<string, FootnoteProperties["numFmt"]> = {
  "1": "decimal",
  a: "lowerLetter",
  A: "upperLetter",
  i: "lowerRoman",
  I: "upperRoman",
} as Record<string, FootnoteProperties["numFmt"]>;

/** Note bodies keyed by reference id, threaded through ODT emission. */
export interface NotesContext {
  footnotes: Map<number, NoteChildren>;
  endnotes: Map<number, NoteChildren>;
  bookmarkNames: Map<number, string>;
  bookmarkEndNames: Map<number, string[]>;
  bookmarkXmlIdCounts: Map<number, number>;
  sequenceNames: Set<string>;
  variableTypes: Map<string, OdtVariableValueType>;
  forms: OfficeFormsOptions;
  revisions: Map<string, CanonicalRevision>;
  comments: Map<number, NonNullable<DocumentOptions["comments"]>[number]>;
  annotationNames: Map<number, string>;
  sectionSources: Map<string, import("./options").OdtSectionSourceOptions>;
}

export function sectionChildText(child: SectionChild): string {
  if (typeof child === "string") return child;
  if ("paragraph" in child)
    return typeof child.paragraph === "string" ? child.paragraph : (child.paragraph.text ?? "");
  return "";
}

/** Allocate a unique XML ID when the canonical bookmark id occurs repeatedly. */
export function nextBookmarkXmlId(id: number, notes: NotesContext): string {
  const occurrence = (notes.bookmarkXmlIdCounts.get(id) ?? 0) + 1;
  notes.bookmarkXmlIdCounts.set(id, occurrence);
  return occurrence === 1 ? `bookmark-${id}` : `bookmark-${id}-${occurrence}`;
}

/** Note reference renders inline as text:note carrying its body paragraphs. */
export function noteXml(
  reference: number | { id: number },
  notes: NotesContext,
  noteClass: "footnote" | "endnote",
  styles: string[],
  images: OdtImage[],
  numbering: DocumentOptions["numbering"],
  charts: OdtChart[],
  embeddedObjects: OdtEmbeddedObjectOptions[] = [],
): string {
  const id = typeof reference === "number" ? reference : reference.id;
  const children = (noteClass === "endnote" ? notes.endnotes : notes.footnotes).get(id) ?? [];
  return xmlElement("text:note", { "text:id": `${noteClass}${id}`, "text:note-class": noteClass }, [
    xmlElement("text:note-citation", undefined, [String(id)]),
    xmlElement("text:note-body", undefined, [
      blocksXml(
        children.map((noteChild) =>
          typeof noteChild === "string" || !("paragraph" in noteChild || "table" in noteChild)
            ? ({ paragraph: noteChild } as SectionChild)
            : (noteChild as SectionChild),
        ),
        styles,
        images,
        notes,
        numbering,
        charts,
        embeddedObjects,
      ),
    ]),
  ]);
}

export function annotationXml(
  author: string | undefined,
  date: Date | string | null | undefined,
  paragraphs: string[],
  name?: string,
): string {
  return xmlElement(
    "office:annotation",
    {
      "office:name": name,
      "dc:creator": author,
      "dc:date": date instanceof Date ? date.toISOString() : (date ?? undefined),
    },
    paragraphs.map((paragraph) => xmlElement("text:p", undefined, [escapeText(paragraph)])),
  );
}

export function bibliographyMarkXml(bibliographyType: string, cachedValue?: string): string {
  return xmlElement(
    "text:bibliography-mark",
    { "text:bibliography-type": bibliographyType },
    cachedValue ? [escapeText(cachedValue)] : [],
  );
}

export function indexXml(
  value: Extract<SectionChild, { toc: unknown }>["toc"],
  styles: string[],
  images: OdtImage[],
  notes: NotesContext,
  numbering: DocumentOptions["numbering"],
  charts: OdtChart[],
  embeddedObjects: OdtEmbeddedObjectOptions[] = [],
): string {
  const entryXml = (entries: SectionChild[] | undefined): string[] =>
    (entries ?? []).map((entry) =>
      blockXml(entry, styles, images, notes, numbering, charts, embeddedObjects),
    );
  const blockXmlGroup = (entries: SectionChild[] | undefined): string =>
    blocksXml(entries ?? [], styles, images, notes, numbering, charts, embeddedObjects);
  if (value.captionLabelIncludingNumbers === undefined) {
    return xmlElement("text:table-of-content", { "text:name": value.alias }, [
      xmlElement("text:table-of-content-source", { "text:outline-level": 10 }),
      indexBodyXml([
        blockXmlGroup(value.leading),
        ...entryXml(value.entries),
        blockXmlGroup(value.trailing),
      ]),
    ]);
  }
  return xmlElement(
    "text:illustration-index",
    {
      "text:name": value.alias,
    },
    [
      xmlElement("text:illustration-index-source", {
        "text:index-scope": "document",
        "text:use-caption": value.captionLabelIncludingNumbers !== undefined,
      }),
      xmlElement("text:index-body", undefined, entryXml(value.entries)),
    ],
  );
}

function indexBodyXml(texts: string[]): string {
  return xmlElement("text:index-body", undefined, texts);
}

export function bibliographyIndexXml(
  value: Extract<SectionChild, { sdt: unknown }>["sdt"],
): string {
  return xmlElement(
    "text:bibliography",
    {
      "text:name": value.properties.alias,
    },
    [
      xmlElement("text:bibliography-source", {}, [
        xmlElement("text:bibliography-entry-template", {
          "text:bibliography-type": "article",
          "text:style-name": "Index",
        }),
      ]),
      xmlElement(
        "text:index-body",
        undefined,
        value.children
          ?.map(sectionChildText)
          .map((text) => xmlElement("text:p", undefined, [escapeText(text)])),
      ),
    ],
  );
}

export function encodeSectionTag(styleName: string): string {
  return `odf:text-section;style=${styleName.replace(/([\\;])/g, "\\$1")}`;
}

export function trackBookmarkId(context: ParseContext, name: string, id: number): void {
  assertSafeBookmarkId(id, "bookmark-id");
  context.bookmarkIdAllocator.value = Math.max(context.bookmarkIdAllocator.value, id);
  context.bookmarkIds.set(name, id);
}

export function nextBookmarkId(context: ParseContext): number {
  const id = context.bookmarkIdAllocator.value + 1;
  assertSafeBookmarkId(id, "bookmark-id");
  context.bookmarkIdAllocator.value = id;
  return id;
}

export interface BookmarkIdDiagnostic {
  elementName?: string;
  path?: string;
}

export function parseBookmarkId(digits: string, diagnostic?: BookmarkIdDiagnostic): number {
  const id = Number(digits);
  assertSafeBookmarkId(id, "bookmark-id", diagnostic);
  return id;
}

export function bookmarkDisplayName(id: number, name: string, notes: NotesContext): string {
  const duplicateCount = bookmarkRawNameCounts.get(notes)?.get(name) ?? 1;
  const counts = duplicateBookmarkCounts.get(notes) ?? new Map<string, number>();
  duplicateBookmarkCounts.set(notes, counts);
  const key = `${id}\0${name}`;
  const occurrence = counts.get(key) ?? 0;
  counts.set(key, occurrence + 1);
  if (duplicateCount < 2) return name;
  return occurrence === 0
    ? `${name}__office-open-id-${id}__`
    : `${name}__office-open-id-${id}-${occurrence}__`;
}

const duplicateBookmarkCounts = new WeakMap<NotesContext, Map<string, number>>();
const bookmarkRawNameCounts = new WeakMap<NotesContext, Map<string, number>>();

export function countBookmarkName(id: number, name: string, notes: NotesContext): void {
  const counts = bookmarkRawNameCounts.get(notes) ?? new Map<string, number>();
  bookmarkRawNameCounts.set(notes, counts);
  counts.set(name, (counts.get(name) ?? 0) + 1);
  notes.bookmarkNames.set(id, name);
}

export function canonicalBookmarkName(
  name: string,
  diagnostic?: BookmarkIdDiagnostic,
): { name: string; encodedId?: number } {
  const match = /^(?<name>.+)__office-open-id-(?<id>\d+)(?:-\d+)?__$/.exec(name);
  if (!match) return { name };
  return {
    name: match.groups?.name ?? name,
    encodedId: parseBookmarkId(match.groups?.id ?? "", diagnostic),
  };
}

function assertSafeBookmarkId(
  id: number,
  fallbackName: string,
  diagnostic?: BookmarkIdDiagnostic,
): void {
  if (!Number.isSafeInteger(id) || id < 0) {
    throw new OdtParseError(
      "ODT bookmark id is not a non-negative safe integer",
      "content.xml",
      diagnostic?.path ?? "/office:body/office:text",
      diagnostic?.elementName ?? fallbackName,
      "unsafe-bookmark-id",
    );
  }
}

export function parseAnnotation(element: Element): ParagraphChild {
  if (attributeString(element, "office:display") === "true") throw unknownOdtElement(element);
  return {
    comment: {
      author: attributeString(element, "dc:creator"),
      date: attributeString(element, "dc:date"),
      children: childrenNamed(element, "text:p").map((paragraph) => ({ text: textOf(paragraph) })),
    },
  };
}

export function parseAnnotationStart(element: Element, context: ParseContext): ParagraphChild[] {
  const name = attributeString(element, "office:name");
  if (!name) return [parseAnnotation(element)];
  const id = context.comments.length + 1;
  context.annotationIds.set(name, id);
  context.comments.push({
    id,
    author: attributeString(element, "dc:creator"),
    date: attributeString(element, "dc:date"),
    children: childrenNamed(element, "text:p").map((paragraph) => textOf(paragraph)),
  });
  return [{ commentRangeStart: { id } }, { commentReference: id }];
}

export function parseAnnotationEnd(element: Element, context: ParseContext): ParagraphChild[] {
  const name = attributeString(element, "office:name") ?? "";
  let id = context.annotationIds.get(name);
  if (id === undefined) {
    if (!name) throw unknownOdtElement(element);
    id = context.comments.length + 1;
    context.annotationIds.set(name, id);
    context.comments.push({ id, children: [] });
  }
  return [{ commentRangeEnd: { id } }];
}

export function parseIndex(
  element: Element,
  context: ParseContext,
): Extract<SectionChild, { toc: unknown }>["toc"] {
  const source = childNamed(element, "text:illustration-index-source");
  const title = childNamed(source, "text:index-title-template");
  const scope = attributeString(source, "text:index-scope");
  if (scope !== undefined && scope !== "document") throw unknownOdtElement(element);
  const body = childNamed(element, "text:index-body");
  const entries = body
    ? parseBlocks(
        body.elements ?? [],
        context,
        0,
        undefined,
        `${context.sourcePath ?? "/office:body/office:text"}/${element.name}/text:index-body`,
      )
    : [];
  return {
    alias: attributeString(element, "text:name"),
    captionLabelIncludingNumbers:
      attributeString(source, "text:use-caption") === "true" ? "Figure" : undefined,
    entries: title ? [{ paragraph: { text: textOf(title) } }, ...entries] : entries,
  };
}

export function parseContentIndex(
  element: Element,
  sourceName: string,
  context: ParseContext,
): Extract<SectionChild, { toc: unknown }>["toc"] {
  const source = childNamed(element, sourceName);
  const title = childNamed(source, "text:index-title-template");
  const body = childNamed(element, "text:index-body");
  const entries = body
    ? parseBlocks(
        body.elements ?? [],
        context,
        0,
        undefined,
        `${context.sourcePath ?? "/office:body/office:text"}/${element.name}/text:index-body`,
      )
    : [];
  return {
    alias: attributeString(element, "text:name"),
    hyperlink: childrenNamed(source, "text:table-of-content-entry-template").some((template) =>
      childNamed(template, "text:index-entry-link-start"),
    ),
    ...(sourceName === "text:table-of-content-source"
      ? {
          headingStyleRange: `1-${attributeNumber(source, "text:outline-level") ?? 10}`,
        }
      : {}),
    entries: title ? [{ paragraph: { text: textOf(title) } }, ...entries] : entries,
  };
}

export function parseBibliography(
  element: Element,
): Extract<SectionChild, { sdt: unknown }>["sdt"] {
  return {
    properties: {
      alias: attributeString(element, "text:name"),
      bibliography: true,
      richText: true,
      ...(attributeString(element, "text:style-name")
        ? { runProperties: { style: attributeString(element, "text:style-name") } }
        : {}),
    },
    children: childrenNamed(element, "text:index-body").flatMap((body) =>
      childrenNamed(body, "text:p").map((paragraph) => ({
        paragraph: { text: textOf(paragraph) },
      })),
    ),
  };
}

export function parseNote(element: Element, context: ParseContext): ParagraphChild[] {
  const body = childNamed(element, "text:note-body");
  const children = body
    ? parseBlocks(
        body.elements ?? [],
        context,
        0,
        undefined,
        `${context.sourcePath ?? "/office:body/office:text"}/${element.name}/text:note-body`,
      )
    : [];
  const isEndnote = attributeString(element, "text:note-class") === "endnote";
  const notes = isEndnote ? context.notes.endnotes : context.notes.footnotes;
  notes.push({ id: notes.length + 1, children });
  const reference = isEndnote
    ? { endnoteReference: notes.length }
    : { footnoteReference: notes.length };
  return [reference];
}
