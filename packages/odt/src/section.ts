import type { FormContainerOptions } from "@office-open/core";
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

import {
  assertElementNames,
  blocksXml,
  parseBlocks,
  unknownOdtElement,
  type ParseContext,
} from "./body";
import type { CanonicalFormControl, OdtChart, OdtImage } from "./drawing";
import type { OdtEmbeddedObjectOptions } from "./package";
import type { CanonicalRevision } from "./paragraph";

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
  sequenceNames: Set<string>;
  variableTypes: Map<string, "float" | "string">;
  forms: FormContainerOptions[];
  revisions: Map<number, CanonicalRevision>;
}

export function sectionChildText(child: SectionChild): string {
  if (typeof child === "string") return child;
  if ("paragraph" in child)
    return typeof child.paragraph === "string" ? child.paragraph : (child.paragraph.text ?? "");
  return "";
}

export function formControlXml(control: CanonicalFormControl): string {
  const common = {
    "form:id": control.id,
    "form:name": control.name,
    "form:disabled": control.disabled,
    "form:tab-index": control.tabIndex,
    "form:automatic-focus": control.automaticFocus,
  };
  if (control.control === "checkBox")
    return xmlElement("form:checkbox", {
      ...common,
      "form:current-state": control.checked ? "checked" : "unchecked",
    });
  if (control.control === "dropDownList")
    return xmlElement(
      "form:listbox",
      common,
      control.entries.map((label, optionIndex) =>
        xmlElement(
          "form:option",
          { "form:selected": control.selectedIndex === optionIndex || undefined },
          [escapeText(label)],
        ),
      ),
    );
  return xmlElement("form:text", {
    ...common,
    "form:current-value": control.value,
    "form:max-length": control.maxLength,
  });
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
): string {
  return xmlElement(
    "office:annotation",
    {
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

export function indexXml(value: Extract<SectionChild, { toc: unknown }>["toc"]): string {
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
      xmlElement(
        "text:index-body",
        undefined,
        value.entries
          ?.map(sectionChildText)
          .map((text) => xmlElement("text:p", undefined, [escapeText(text)])),
      ),
    ],
  );
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

export function nextBookmarkId(context: ParseContext): number {
  return Math.max(0, ...context.bookmarkIds.values()) + 1;
}

export function parseForms(element: Element, context: ParseContext): SectionChild[] {
  const forms = childrenNamed(element, "form:form");
  const parsedForms = forms.map((form) => {
    assertElementNames(form, ["form:name"], ["form:text", "form:checkbox", "form:listbox"]);
    return {
      ...(attributeString(form, "form:name") ? { name: attributeString(form, "form:name") } : {}),
      ...(attributeString(element, "form:automatic-focus") === "true"
        ? { automaticFocus: true }
        : {}),
      ...(attributeString(element, "form:apply-design-mode") === "true"
        ? { designMode: true }
        : {}),
      controls:
        form.elements
          ?.filter((child) => child.type === "element")
          .map((control) => parseFormControl(control)) ?? [],
    };
  });
  context.notes.forms.push(...parsedForms);
  return [];
}

export function parseFormControl(control: Element): CanonicalFormControl {
  const id = attributeString(control, "form:id");
  const name = attributeString(control, "form:name");
  if (id && !/^field\d+$/.test(id)) {
    if (name) throw unknownOdtElement(control);
    if (!/^[A-Za-z_:][\w.:-]*$/.test(id)) throw unknownOdtElement(control);
  }
  const common = {
    ...((name ?? id) ? { name: name ?? id } : {}),
    ...(attributeString(control, "form:disabled") === "true" ? { disabled: true } : {}),
    ...(attributeNumber(control, "form:tab-index") !== undefined
      ? { tabIndex: attributeNumber(control, "form:tab-index") }
      : {}),
  };
  return (() => {
    if (control.name === "form:text") {
      assertElementNames(
        control,
        [
          "form:id",
          "form:name",
          "form:disabled",
          "form:tab-index",
          "form:current-value",
          "form:max-length",
        ],
        [],
      );
      const value = attributeString(control, "form:current-value");
      return {
        control: "text",
        ...common,
        ...(value ? { value } : {}),
        ...(attributeNumber(control, "form:max-length") !== undefined
          ? { maxLength: attributeNumber(control, "form:max-length") }
          : {}),
      };
    }
    if (control.name === "form:checkbox") {
      assertElementNames(
        control,
        ["form:id", "form:name", "form:disabled", "form:tab-index", "form:current-state"],
        [],
      );
      const state = attributeString(control, "form:current-state");
      if (state !== "checked" && state !== "unchecked") throw unknownOdtElement(control);
      return {
        control: "checkBox",
        ...common,
        checked: state === "checked",
      };
    }
    assertElementNames(
      control,
      ["form:id", "form:name", "form:disabled", "form:tab-index", "form:automatic-focus"],
      ["form:option", "form:item"],
    );
    const entries = (control.elements ?? [])
      .filter((child) => ["form:option", "form:item"].includes(child.name ?? ""))
      .map((child) => {
        assertElementNames(child, ["form:selected"]);
        return textOf(child);
      });
    const selectedIndex = (control.elements ?? []).findIndex(
      (child) => attributeString(child, "form:selected") === "true",
    );
    return {
      control: "dropDownList",
      ...common,
      entries,
      ...(selectedIndex >= 0 ? { selectedIndex } : {}),
    };
  })();
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

export function parseIndex(element: Element): Extract<SectionChild, { toc: unknown }>["toc"] {
  const source = childNamed(element, "text:illustration-index-source");
  const title = childNamed(source, "text:index-title-template");
  const scope = attributeString(source, "text:index-scope");
  if (scope !== undefined && scope !== "document") throw unknownOdtElement(element);
  const entries = childrenNamed(childNamed(element, "text:index-body"), "text:p").map(
    (paragraph): SectionChild => ({ paragraph: { text: textOf(paragraph) } }),
  );
  return {
    alias: attributeString(element, "text:name"),
    captionLabelIncludingNumbers:
      attributeString(source, "text:use-caption") === "true" ? "Figure" : undefined,
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
  const children = body ? parseBlocks(body.elements ?? [], context) : [];
  const isEndnote = attributeString(element, "text:note-class") === "endnote";
  const notes = isEndnote ? context.notes.endnotes : context.notes.footnotes;
  notes.push({ id: notes.length + 1, children });
  const reference = isEndnote
    ? { endnoteReference: notes.length }
    : { footnoteReference: notes.length };
  return [reference];
}
