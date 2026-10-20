import type {
  FormControlOptions,
  SequenceDeclarationOptions,
  VariableDeclarationOptions,
} from "@office-open/core";
import type {
  ChartOptions,
  DocumentOptions,
  FormFieldOptions,
  ParagraphChild,
  ParagraphOptions,
  PictureOptions,
  RunOptions,
  SectionChild,
  ShapeOptions,
} from "@office-open/docx";
import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  escapeText,
  textOf,
  officeFormsXml,
  xmlElement,
} from "@office-open/odf";
import type { Element } from "@office-open/xml";

import { addCharacterStyle, addParagraphStyle, characterProperties } from "./automatic-styles";
import {
  assertElementNames,
  malformedOdtElement,
  unknownOdtElement,
  type ParseContext,
} from "./body";
import {
  chartFrameXml,
  isShapeElement,
  parseGroup,
  parseChartFrame,
  parseCustomShape,
  parsePictureFrame,
  parseShape,
  parseShapeHyperlink,
  parseTextBoxFrame,
  pictureFrameXml,
  wpsShapeFrameXml,
  type OdtChart,
  type OdtImage,
} from "./drawing";
import type { OdtEmbeddedObjectOptions } from "./embedded-objects";
import { OdtParseError } from "./error";
import { lengthToTwips, twipsToLength } from "./master-pages";
import {
  annotationXml,
  bibliographyMarkXml,
  nextBookmarkId,
  noteXml,
  parseAnnotationEnd,
  parseAnnotationStart,
  parseNote,
  type NotesContext,
} from "./section";

/** Track-change types derived from the canonical paragraph child union. */
export type CanonicalTrackChange = Extract<ParagraphChild, { insertion: unknown }>["insertion"];

export type ChangedProperties = Omit<CanonicalTrackChange, "children">;

export type TrackChangeChild = CanonicalTrackChange["children"][number];

export type CanonicalSimpleField = Extract<ParagraphChild, { simpleField: unknown }>["simpleField"];
type RunChild = NonNullable<RunOptions["children"]>[number];
type CanonicalRuby = Extract<RunChild, { ruby: unknown }>["ruby"];

/** Canonical tracked change metadata and deleted content for ODF emission. */
export type CanonicalRevision = ChangedProperties & {
  kind?: "format";
  children?: TrackChangeChild[];
};

export type ParsedTrackChange = ChangedProperties & {
  kind: "insertion" | "deletion" | "format";
  children?: Element[];
};

const SENDER_FIELD_SUFFIXES: Record<string, string> = {
  "text:sender-firstname": "firstname",
  "text:sender-lastname": "lastname",
  "text:sender-initials": "initials",
  "text:sender-title": "title",
  "text:sender-position": "position",
  "text:sender-email": "email",
  "text:sender-phone-private": "phone-private",
  "text:sender-fax": "fax",
  "text:sender-company": "company",
  "text:sender-phone-work": "phone-work",
  "text:sender-street": "street",
  "text:sender-city": "city",
  "text:sender-postal-code": "postal-code",
  "text:sender-country": "country",
  "text:sender-state-or-province": "state-or-province",
};

const SENDER_FIELD_ELEMENTS = new Map(
  Object.entries(SENDER_FIELD_SUFFIXES).map(([name, suffix]) => [`AUTHOR \\${suffix}`, name]),
);

/** Docx tab stop, indexed from the shared paragraph model. */
export type TabStop = NonNullable<ParagraphOptions["tabStops"]>[number];

export function isTrackChangeChild(child: string | ParagraphChild): child is TrackChangeChild {
  if (typeof child === "string") return true;
  return (
    "text" in child ||
    "break" in child ||
    "pageBreak" in child ||
    "columnBreak" in child ||
    "picture" in child ||
    "chart" in child ||
    "wpsShape" in child ||
    "formField" in child
  );
}

function isRunChild(child: TrackChangeChild): child is RunOptions {
  return (
    typeof child !== "string" &&
    "text" in child &&
    !("hyperlink" in child) &&
    !("insertion" in child) &&
    !("deletion" in child)
  );
}

function formatChangeChildren(
  children: readonly (string | ParagraphChild)[],
  revision: ChangedProperties,
): TrackChangeChild[] {
  return children.map((child): TrackChangeChild => {
    if (typeof child === "string") return { text: child, revision };
    if (isTrackChangeChild(child) && isRunChild(child)) return { ...child, revision };
    throw malformedOdtElement(
      "text:format-change",
      "changed content cannot carry a formatting revision",
    );
  });
}

function nextTrackChangeId(context: ParseContext): number {
  const changes = [...context.changes.values(), ...context.changeAlternatives.values()];
  return Math.max(0, ...changes.map((change) => change.id ?? 0)) + 1;
}

/** Runs of two or more spaces emit text:s so XML whitespace folding keeps them. */
export function spacesXml(text: string): string {
  const parts: string[] = [];
  let index = 0;
  for (const match of text.matchAll(/ {2,}/g)) {
    const start = match.index ?? 0;
    parts.push(escapeText(text.slice(index, start)));
    parts.push(xmlElement("text:s", { "text:c": match[0].length }));
    index = start + match[0].length;
  }
  parts.push(escapeText(text.slice(index)));
  return parts.join("");
}

/** ODF style:tab-stop → the closest typed docx tab stop. */
export function parseTabStop(element: Element): TabStop {
  const type = attributeString(element, "style:type");
  const leaderText = attributeString(element, "style:leader-text");
  return {
    type: type === "char" ? "decimal" : ((type ?? "left") as TabStop["type"]),
    position: lengthToTwips(attributeString(element, "style:position")) ?? 0,
    leader:
      leaderText === "."
        ? "dot"
        : leaderText === "-"
          ? "hyphen"
          : leaderText === "_"
            ? "underscore"
            : leaderText === "·"
              ? "middleDot"
              : undefined,
  };
}

export function decodeStyleName(value: string): string {
  return value.replace(/\\([\\;])/g, "$1");
}

export function normalizeParagraph(input: string | ParagraphOptions): ParagraphOptions {
  return typeof input === "string" ? { text: input } : input;
}

export function paragraphXml(
  options: ParagraphOptions,
  styles: string[],
  images: OdtImage[],
  notes: NotesContext,
  numbering: DocumentOptions["numbering"],
  charts: OdtChart[],
  embeddedObjects: OdtEmbeddedObjectOptions[] = [],
): string {
  const children = runXml(options, styles, images, notes, numbering, charts, embeddedObjects);
  if (
    options.children?.length &&
    options.children.every((child) => typeof child === "object" && "formField" in child)
  ) {
    return "";
  }
  const alignment = typeof options.alignment === "string" ? options.alignment : undefined;
  const styleName =
    options.style ??
    (alignment || options.pageBreakBefore || options.tabStops?.length
      ? addParagraphStyle(
          { alignment, pageBreakBefore: options.pageBreakBefore, tabStops: options.tabStops },
          styles,
        )
      : undefined);
  const heading = /^Heading([1-9])$/.exec(options.heading ?? "");
  const attributes = {
    "text:style-name": styleName,
    "text:outline-level": heading ? Number(heading[1]) : undefined,
  };
  return xmlElement(heading ? "text:h" : "text:p", attributes, children);
}

export function runXml(
  options: ParagraphOptions,
  styles: string[],
  images: OdtImage[],
  notes: NotesContext,
  numbering: DocumentOptions["numbering"],
  charts: OdtChart[],
  embeddedObjects: OdtEmbeddedObjectOptions[] = [],
): string[] {
  if (options.text !== undefined && options.children === undefined) {
    return [spacesXml(options.text)];
  }
  return (options.children ?? []).map((child) => {
    if (typeof child === "string") return spacesXml(child);
    if ("pageBreak" in child) return "<text:soft-page-break/>";
    if ("columnBreak" in child) return "<text:line-break/>";
    if ("break" in child) return lineBreakXml(child as RunOptions);
    if ("bookmark" in child) {
      return xmlElement("text:bookmark", {
        "text:name": (child as { bookmark: { name: string } }).bookmark.name,
      });
    }
    if ("bookmarkStart" in child) {
      return xmlElement("text:bookmark-start", {
        "xml:id": child.bookmarkStart.id ? `bookmark-${child.bookmarkStart.id}` : undefined,
        "text:name": child.bookmarkStart.name,
      });
    }
    if ("bookmarkEnd" in child) {
      return bookmarkEndXml(child.bookmarkEnd.id, notes);
    }
    if ("simpleField" in child) {
      return simpleFieldXml(child.simpleField, notes);
    }
    if ("comment" in child) {
      return annotationXml(
        child.comment.author,
        child.comment.date,
        child.comment.children.map((commentChild) => commentChildText(commentChild)),
      );
    }
    if ("commentRangeStart" in child) {
      const id = child.commentRangeStart.id;
      const comment = notes.comments.get(id);
      if (!comment) throw orphanCommentRange(id);
      notes.annotationNames.set(id, `__Annotation__${id}__office-open`);
      return annotationXml(
        comment.author,
        comment.date,
        comment.children.map((commentChild) => commentChildText(commentChild)),
        notes.annotationNames.get(id),
      );
    }
    if ("commentReference" in child) return "";
    if ("commentRangeEnd" in child) {
      const id = child.commentRangeEnd.id;
      const name = notes.annotationNames.get(id);
      if (!name) throw orphanCommentRange(id);
      return xmlElement("office:annotation-end", { "office:name": name });
    }
    if ("formField" in child) return formFieldXml(child.formField, notes);
    if ("insertion" in child) {
      return insertionXml(
        child.insertion,
        notes,
        styles,
        images,
        numbering,
        charts,
        embeddedObjects,
      );
    }
    if ("deletion" in child) {
      const id = child.deletion.id ?? notes.revisions.size + 1;
      notes.revisions.set(id, {
        id,
        author: child.deletion.author,
        date: child.deletion.date,
        children: child.deletion.children,
      });
      return xmlElement("text:change", { "text:change-id": `rev${id}` });
    }
    if ("footnoteReference" in child) {
      return noteXml(
        (child as { footnoteReference: number | { id: number } }).footnoteReference,
        notes,
        "footnote",
        styles,
        images,
        numbering,
        charts,
        embeddedObjects,
      );
    }
    if ("endnoteReference" in child) {
      return noteXml(
        (child as { endnoteReference: number | { id: number } }).endnoteReference,
        notes,
        "endnote",
        styles,
        images,
        numbering,
        charts,
        embeddedObjects,
      );
    }
    if ("hyperlink" in child) {
      const link = (
        child as {
          hyperlink: {
            url?: string;
            anchor?: string;
            tooltip?: string;
            targetFrame?: string;
            children?: Array<string | { text?: string }>;
          };
          text?: string;
        }
      ).hyperlink;
      const href = link.url ?? (link.anchor !== undefined ? `#${link.anchor}` : undefined);
      if (href === undefined) return "";
      const inner = link.children
        ? link.children
            .map((c) => (typeof c === "string" ? spacesXml(c) : spacesXml(c.text ?? "")))
            .join("")
        : spacesXml((child as { text?: string }).text ?? "");
      return xmlElement(
        "text:a",
        {
          "xlink:type": "simple",
          "xlink:href": href,
          "office:title": link.tooltip,
          "office:target-frame-name": link.targetFrame,
        },
        [inner],
      );
    }
    if ("text" in child) {
      const run = child as RunOptions;
      const ruby = run.children?.find(
        (rubyChild): rubyChild is Extract<RunChild, { ruby: unknown }> =>
          typeof rubyChild === "object" && "ruby" in rubyChild,
      )?.ruby;
      if (ruby) {
        return rubyXml(ruby, styles, images, notes, numbering, charts, embeddedObjects);
      }
      const styleName = run.style ?? addCharacterStyle(characterProperties(run), styles);
      const runXml =
        lineBreakXml(run) +
        xmlElement("text:span", { "text:style-name": styleName }, [spacesXml(run.text ?? "")]);
      if (!run.revision) return runXml;
      const id = run.revision.id ?? notes.revisions.size + 1;
      notes.revisions.set(id, {
        id,
        author: run.revision.author,
        date: run.revision.date,
        kind: "format",
      });
      return (
        xmlElement("text:change-start", { "text:change-id": `rev${id}` }) +
        runXml +
        xmlElement("text:change-end", { "text:change-id": `rev${id}` })
      );
    }
    if ("ruby" in child) {
      const ruby = (child as { ruby: CanonicalRuby }).ruby;
      return rubyXml(ruby, styles, images, notes, numbering, charts, embeddedObjects);
    }
    if ("tab" in child) return "<text:tab/>";
    if ("picture" in child)
      return pictureFrameXml((child as { picture: PictureOptions }).picture, images);
    if ("wpsShape" in child)
      return wpsShapeFrameXml(
        (child as { wpsShape: ShapeOptions }).wpsShape,
        styles,
        embeddedObjects,
      );
    if ("chart" in child) {
      return chartFrameXml((child as { chart: ChartOptions }).chart, charts);
    }
    return "";
  });
}

export function simpleFieldXml(field: CanonicalSimpleField, notes: NotesContext): string {
  const instruction = field.instruction.trim();
  const argument = decodeFieldToken(instruction, instruction.indexOf(" ") + 1);
  const remainder = instruction.slice(argument.next).trim();
  if (instruction === "PAGE")
    return xmlElement("text:page-number", undefined, [escapeText(field.cachedValue ?? "")]);
  if (instruction === "NUMPAGES")
    return xmlElement("text:page-count", undefined, [escapeText(field.cachedValue ?? "")]);
  if (instruction === "AUTHOR")
    return xmlElement("text:author-name", undefined, [escapeText(field.cachedValue ?? "")]);
  const senderField = SENDER_FIELD_ELEMENTS.get(instruction);
  if (senderField) return xmlElement(senderField, undefined, [escapeText(field.cachedValue ?? "")]);
  if (instruction.startsWith("DATE"))
    return xmlElement(
      "text:date",
      {
        "style:data-style-name": remainder.startsWith("\\@ ")
          ? decodeFieldArgument(remainder.slice(3))
          : undefined,
      },
      [escapeText(field.cachedValue ?? "")],
    );
  if (instruction.startsWith("ODF USERFIELD "))
    return xmlElement("text:user-field-get", { "text:name": argument.value }, [
      escapeText(field.cachedValue ?? ""),
    ]);
  if (instruction.startsWith("ODF VARIABLEGET "))
    return xmlElement("text:variable-get", { "text:name": argument.value }, [
      escapeText(field.cachedValue ?? ""),
    ]);
  if (instruction.startsWith("ODF SEQUENCEREF "))
    return xmlElement("text:sequence-ref", { "text:ref-name": argument.value }, [
      escapeText(field.cachedValue ?? ""),
    ]);
  if (instruction.startsWith("MERGEFIELD "))
    return xmlElement(
      "text:database-display",
      { "text:database-name": "Database", "text:column-name": argument.value },
      [escapeText(field.cachedValue ?? "")],
    );
  if (instruction.startsWith("DOCPROPERTY ")) {
    if (argument.value === "Title")
      return xmlElement("text:title", undefined, [escapeText(field.cachedValue ?? "")]);
    return xmlElement("text:user-defined", { "text:name": argument.value }, [
      escapeText(field.cachedValue ?? ""),
    ]);
  }
  if (instruction.startsWith("ODF PLACEHOLDER ")) {
    const separator = remainder.indexOf(" ");
    return xmlElement(
      "text:placeholder",
      {
        "text:placeholder-type": separator >= 0 ? argument.value : decodeFieldArgument(remainder),
        "text:description":
          separator >= 0 ? decodeFieldArgument(remainder.slice(separator + 1)) : undefined,
      },
      [escapeText(field.cachedValue ?? "")],
    );
  }
  if (instruction === "FILLIN")
    return xmlElement("text:text-input", undefined, [escapeText(field.cachedValue ?? "")]);
  if (instruction.startsWith("ODF HIDDENPARAGRAPH "))
    return xmlElement("text:hidden-paragraph", { "text:condition": argument.value });
  if (instruction.startsWith("ODF META "))
    return xmlElement("text:meta", { "text:name": argument.value }, [
      escapeText(field.cachedValue ?? ""),
    ]);
  if (instruction.startsWith("REF ")) {
    if (odfReferenceFormat(remainder) === undefined && remainder !== "") {
      throw unsupportedField(instruction);
    }
    return bookmarkReferenceXml({
      name: argument.value,
      referenceFormat: odfReferenceFormat(remainder),
    });
  }
  let formula: string | undefined;
  if (instruction.startsWith("SEQ ")) {
    if (remainder && !remainder.startsWith("=")) throw unsupportedField(instruction);
    if (remainder) formula = remainder.slice(1).trim();
    notes.sequenceNames.add(argument.value);
    return sequenceXml({
      name: argument.value,
      formula,
      display: field.cachedValue,
    });
  }
  if (instruction.startsWith("STYLEREF ")) {
    const match = /^STYLEREF\s+(?<level>[1-9])\s+\\n$/.exec(instruction);
    if (!match) throw unsupportedField(instruction);
    return chapterXml({
      display: field.cachedValue,
      outlineLevel: Number(match.groups?.level),
    });
  }
  if (instruction.startsWith("VARIABLE ")) {
    if (remainder && !remainder.startsWith("=")) throw unsupportedField(instruction);
    if (remainder) formula = remainder.slice(1).trim();
    const valueType = Number.isFinite(Number(field.cachedValue)) ? "float" : "string";
    notes.variableTypes.set(argument.value, valueType);
    return variableSetXml({
      name: argument.value,
      valueType,
      value: field.cachedValue,
      display: field.cachedValue,
      formula,
    });
  }
  if (instruction.startsWith("IF ")) {
    const match = /^"(?<content>(?:[^\\"]|\\[\s\S])*)"\s+""$/.exec(remainder);
    if (!match) throw unsupportedField(instruction);
    return hiddenTextXml({
      condition: argument.value,
      content: decodeFieldArgument(match.groups?.content ?? ""),
      hidden: field.cachedValue === undefined ? true : undefined,
    });
  }
  if (instruction.startsWith("=")) {
    return expressionXml({
      formula: instruction.slice(1).trim(),
      valueType: Number.isFinite(Number(field.cachedValue)) ? "float" : "string",
      value: field.cachedValue,
      display: field.cachedValue,
    });
  }
  if (instruction.startsWith("CITATION ")) {
    if (remainder) throw unsupportedField(instruction);
    return bibliographyMarkXml(argument.value, field.cachedValue);
  }
  throw unsupportedField(instruction);
}

export function decodeFieldArgument(value: string): string {
  return value.replace(/\\([\\"])/g, "$1");
}

export function decodeFieldToken(
  instruction: string,
  start: number,
): { value: string; next: number } {
  let index = start;
  let value = "";
  while (index < instruction.length) {
    const character = instruction[index];
    if (character === undefined) break;
    if (character === "\\" && index + 1 < instruction.length) {
      const next = instruction[index + 1]!;
      if (next === "\\" || next === '"' || /\s/.test(next)) {
        value += next;
        index += 2;
        continue;
      }
      break;
    }
    if (/\s/.test(character)) break;
    value += character;
    index += 1;
  }
  return { value, next: index };
}

export function unsupportedField(instruction: string): OdtParseError {
  return new OdtParseError(
    `content.xml: field instruction has no ODT mapping: ${instruction}`,
    "content.xml",
    "/office:document-content/office:body/office:text",
    "text:field",
    "unsupported field instruction",
  );
}

export function orphanBookmarkEnd(id: number): OdtParseError {
  return new OdtParseError(
    `content.xml: bookmark end ${id} has no canonical bookmark start`,
    "content.xml",
    "/office:document-content/office:body/office:text/text:bookmark-end",
    "text:bookmark-end",
    "no canonical bookmark start",
  );
}

export function orphanCommentRange(id: number): OdtParseError {
  return new OdtParseError(
    `content.xml: comment range ${id} has no canonical comment body`,
    "content.xml",
    "/office:document-content/office:body/office:text/office:annotation",
    "office:annotation",
    "no canonical comment body",
  );
}

function commentChildText(
  child:
    | string
    | ParagraphOptions
    | SectionChild
    | NonNullable<Extract<ParagraphChild, { comment: unknown }>["comment"]>["children"][number],
): string {
  if (typeof child === "string") return child;
  if ("paragraph" in child) {
    if (typeof child.paragraph === "string") return child.paragraph;
    return child.paragraph.text ?? "";
  }
  if ("text" in child && typeof child.text === "string") return child.text;
  return "";
}

export function odfReferenceFormat(fieldSwitch: string | undefined): string | undefined {
  if (fieldSwitch === "\\p") return "page";
  if (fieldSwitch === "\\r") return "number";
  return undefined;
}

export function fieldReferenceSwitch(format: string | undefined): string {
  if (format === "page") return " \\p";
  if (format === "number") return " \\r";
  return "";
}

export function formFieldXml(field: FormFieldOptions, notes: NotesContext): string {
  const common = {
    ...(field.name !== undefined ? { name: field.name } : {}),
    ...(field.enabled === false ? { disabled: true } : {}),
    ...(field.tabIndex !== undefined ? { tabIndex: field.tabIndex } : {}),
  };
  const control: FormControlOptions = field.checkBox
    ? {
        control: "checkbox",
        ...common,
        currentState:
          (field.checkBox.checked ?? field.checkBox.default ?? false) ? "checked" : "unchecked",
      }
    : field.dropDownList
      ? {
          control: "listBox",
          ...common,
          options: field.dropDownList.entries.map((label, index) => ({
            label,
            currentSelected: index === field.dropDownList?.result,
          })),
        }
      : {
          control: "text",
          ...common,
          ...(field.textInput?.value !== undefined ? { currentValue: field.textInput.value } : {}),
          ...(field.textInput?.maxLength !== undefined
            ? { maxLength: field.textInput.maxLength }
            : {}),
        };
  notes.forms.containers.push({ controls: [control] });
  return "";
}

export function insertionXml(
  value: ChangedProperties & { children?: TrackChangeChild[] },
  notes: NotesContext,
  styles: string[],
  images: OdtImage[],
  numbering: DocumentOptions["numbering"],
  charts: OdtChart[],
  embeddedObjects: OdtEmbeddedObjectOptions[] = [],
): string {
  const id = value.id ?? notes.revisions.size + 1;
  notes.revisions.set(id, { id, author: value.author, date: value.date });
  const children = runXml(
    { children: value.children },
    styles,
    images,
    notes,
    numbering,
    charts,
    embeddedObjects,
  );
  return (
    xmlElement("text:change-start", { "text:change-id": `rev${id}` }) +
    children.join("") +
    xmlElement("text:change-end", { "text:change-id": `rev${id}` })
  );
}

export function declarationsXml(options: DocumentOptions, notes: NotesContext): string {
  const implicitSequenceDeclarations = [...notes.sequenceNames]
    .filter(
      (name) => !options.sequenceDeclarations?.some((declaration) => declaration.name === name),
    )
    .map((name): SequenceDeclarationOptions => ({ name }));
  const sequenceDecls =
    options.sequenceDeclarations?.length || implicitSequenceDeclarations.length
      ? xmlElement(
          "text:sequence-decls",
          undefined,
          [...(options.sequenceDeclarations ?? []), ...implicitSequenceDeclarations].map(
            (declaration) =>
              xmlElement("text:sequence-decl", {
                "text:name": declaration.name,
                "text:display-outline-level": declaration.displayOutlineLevel,
                "text:outline-level-separator": declaration.outlineSeparator,
              }),
          ),
        )
      : "";
  const implicitVariableDeclarations = [...notes.variableTypes]
    .filter(
      ([name]) => !options.variableDeclarations?.some((declaration) => declaration.name === name),
    )
    .map(([name, valueType]): VariableDeclarationOptions => ({ name, valueType }));
  const variableDecls =
    options.variableDeclarations?.length || implicitVariableDeclarations.length
      ? xmlElement(
          "text:variable-decls",
          undefined,
          [...(options.variableDeclarations ?? []), ...implicitVariableDeclarations].map(
            (declaration) =>
              xmlElement("text:variable-decl", {
                "text:name": declaration.name,
                "office:value-type": declaration.valueType,
              }),
          ),
        )
      : "";
  const userFieldNames =
    options.settings?.docVars?.filter((variable) => !variable.name.startsWith("DDE ")) ?? [];
  const ddeConnections =
    options.settings?.docVars?.filter((variable) => variable.name.startsWith("DDE ")) ?? [];
  const userFieldDecls = userFieldNames.length
    ? xmlElement(
        "text:user-field-decls",
        undefined,
        userFieldNames.map((variable) => {
          const numeric = variable.val !== "" && Number.isFinite(Number(variable.val));
          return xmlElement("text:user-field-decl", {
            "text:name": variable.name,
            "office:value-type": numeric ? "float" : "string",
            [numeric ? "office:value" : "office:string-value"]: variable.val,
          });
        }),
      )
    : "";
  const ddeDecls = ddeConnections.length
    ? xmlElement(
        "text:dde-connection-decls",
        undefined,
        ddeConnections.map((variable) => {
          const [application, topic, item] = variable.val.split("|");
          return xmlElement("text:dde-connection-decl", {
            "office:name": variable.name.slice(4),
            "office:dde-application": application,
            "office:dde-topic": topic,
            "office:dde-item": item,
          });
        }),
      )
    : "";
  const revisions = notes.revisions.size
    ? xmlElement(
        "text:tracked-changes",
        undefined,
        [...notes.revisions.entries()].map(([, revision]) =>
          xmlElement("text:changed-region", { "text:id": `rev${revision.id}` }, [
            xmlElement(
              `text:${
                revision.kind === "format"
                  ? "format-change"
                  : revision.children
                    ? "deletion"
                    : "insertion"
              }`,
              undefined,
              [
                xmlElement("office:change-info", undefined, [
                  xmlElement("dc:creator", undefined, [escapeText(revision.author)]),
                  xmlElement("dc:date", undefined, [String(revision.date)]),
                ]),
                ...(revision.children ?? []).map((child) =>
                  xmlElement(
                    "text:p",
                    undefined,
                    runXml({ children: [child] }, [], [], notes, undefined, []),
                  ),
                ),
              ],
            ),
          ]),
        ),
      )
    : "";
  const formsXml = officeFormsXml(notes.forms);
  return formsXml + sequenceDecls + variableDecls + userFieldDecls + ddeDecls + revisions;
}

export function variableSetXml(value: {
  name: string;
  valueType: "float" | "string";
  value?: string;
  display?: string;
  formula?: string;
}): string {
  return xmlElement(
    "text:variable-set",
    {
      "text:name": value.name,
      "office:value-type": value.valueType,
      "office:value": value.value,
      "text:display": value.display,
      "text:formula": value.formula,
    },
    value.display === undefined ? [] : [escapeText(value.display)],
  );
}

export function bookmarkReferenceXml(value: { name: string; referenceFormat?: string }): string {
  return xmlElement(
    "text:bookmark-ref",
    {
      "text:reference-format": value.referenceFormat,
    },
    [escapeText(value.name)],
  );
}

export function bookmarkEndXml(id: number, notes: NotesContext): string {
  const name = notes.bookmarkNames.get(id);
  if (!name) throw orphanBookmarkEnd(id);
  return xmlElement("text:bookmark-end", { "text:name": name });
}

export function chapterXml(value: { display?: string; outlineLevel?: number }): string {
  return xmlElement(
    "text:chapter",
    {
      "text:display": value.display,
      "text:outline-level": value.outlineLevel,
    },
    value.display === undefined ? [] : [escapeText(value.display)],
  );
}

export function expressionXml(value: {
  formula?: string;
  valueType?: "float" | "string";
  value?: string;
  display?: string;
}): string {
  return xmlElement(
    "text:expression",
    {
      "text:formula": value.formula,
      "office:value-type": value.valueType,
      "office:value": value.value,
      "text:display": value.display,
    },
    value.display === undefined ? [] : [escapeText(value.display)],
  );
}

export function sequenceXml(value: {
  name: string;
  referenceName?: string;
  display?: string;
  formula?: string;
}): string {
  return xmlElement(
    "text:sequence",
    {
      "text:name": value.name,
      "text:ref-name": value.referenceName,
      "text:display": value.display,
      "text:formula": value.formula,
    },
    value.display === undefined ? [] : [escapeText(value.display)],
  );
}

export function hiddenTextXml(value: {
  condition: string;
  content: string;
  hidden?: boolean;
  fixed?: boolean;
}): string {
  return xmlElement(
    "text:hidden-text",
    {
      "text:condition": value.condition,
      "text:is-hidden": value.hidden,
      "text:is-fixed": value.fixed,
    },
    [escapeText(value.content)],
  );
}

/** Run breaks render as text:line-break (ODF has no w:br/@clear equivalent). */
export function lineBreakXml(run: RunOptions): string {
  if (!run.break) return "";
  const count = typeof run.break === "number" ? run.break : (run.break.count ?? 1);
  return "<text:line-break/>".repeat(Math.max(0, count));
}

/** Tab stop → ODF style:tab-stop; decimal maps to char with a dot. */
export function tabStopXml(tab: TabStop): string {
  if (
    tab.type !== "left" &&
    tab.type !== "center" &&
    tab.type !== "right" &&
    tab.type !== "decimal"
  )
    return "";
  return xmlElement("style:tab-stop", {
    "style:position": twipsToLength(tab.position),
    "style:type": tab.type === "left" ? undefined : tab.type === "decimal" ? "char" : tab.type,
    "style:char": tab.type === "decimal" ? "." : undefined,
    "style:leader-text":
      tab.leader === "dot"
        ? "."
        : tab.leader === "hyphen"
          ? "-"
          : tab.leader === "underscore"
            ? "_"
            : tab.leader === "middleDot"
              ? "·"
              : undefined,
  });
}

export function withPageBreakBefore(child: SectionChild): SectionChild {
  if (!("paragraph" in child) || typeof child.paragraph === "string") return child;
  return { paragraph: { ...child.paragraph, pageBreakBefore: true } };
}

export function parseParagraph(element: Element, context: ParseContext): ParagraphOptions {
  const style = context.styles.get(attributeString(element, "text:style-name") ?? "");
  const headingLevel = attributeNumber(element, "text:outline-level");
  const runs = parseRuns(element, context);
  const result: ParagraphOptions = {};
  if (style?.alignment) result.alignment = style.alignment as ParagraphOptions["alignment"];
  if (style?.properties?.some((property) => Object.keys(property.attributes).length > 0))
    result.style = attributeString(element, "text:style-name");
  if (style?.pageBreakBefore) result.pageBreakBefore = true;
  if (style?.tabStops) result.tabStops = style.tabStops;
  if (headingLevel && headingLevel <= 6) {
    result.heading = `Heading${headingLevel}` as ParagraphOptions["heading"];
  }
  if (result.heading && context.outline) {
    const outlineLevel = Number(/Heading([1-9])$/.exec(result.heading)?.[1]);
    if (outlineLevel && context.outline.levels.some((level) => level.level === outlineLevel - 1))
      result.numbering = { reference: context.outline.reference, level: outlineLevel - 1 };
  }
  if (runs.length === 1 && typeof runs[0] === "string") result.text = runs[0];
  else if (runs.length > 0) result.children = runs as ParagraphOptions["children"];
  return result;
}

export function parseRuns(
  element: Element,
  context: ParseContext,
): (string | RunOptions | ParagraphChild)[] {
  const nodes = element.elements ?? [];
  let resumeAfter = -1;
  const runs = (element.elements ?? []).flatMap(
    (child, index): (string | RunOptions | ParagraphChild)[] => {
      if (index <= resumeAfter) return [];
      if (child.type === "text") return [String(child.text ?? "")];
      if (child.name === "text:line-break") return [{ break: 1 }];
      if (child.name === "field:fieldmark") return [parseFieldmark(child, "", context)];
      if (child.name === "field:fieldmark-start") {
        const endIndex = nodes.findIndex(
          (node, nodeIndex) => nodeIndex > index && node.name === "field:fieldmark-end",
        );
        if (endIndex < 0) {
          throw malformedOdtElement(child.name, "fieldmark end is missing");
        }
        const separatorIndex = nodes.findIndex(
          (node, nodeIndex) =>
            nodeIndex > index && nodeIndex < endIndex && node.name === "field:fieldmark-separator",
        );
        const valueElement: Element = {
          name: "field:value",
          elements: nodes.slice((separatorIndex >= 0 ? separatorIndex : index) + 1, endIndex),
        };
        const parsed = parseFieldmark(child, textOf(valueElement), context);
        resumeAfter = endIndex;
        return [parsed];
      }
      if (child.name === "text:soft-page-break") return [{ pageBreak: true }];
      if (child.name === "text:s") {
        const count = attributeNumber(child, "text:c") ?? 1;
        return [" ".repeat(count)];
      }
      if (child.name === "text:a") {
        const href = attributeString(child, "xlink:href") ?? "";
        return [
          {
            hyperlink: {
              ...(href.startsWith("#") ? { anchor: href.slice(1) } : { url: href }),
              tooltip: attributeString(child, "office:title"),
              targetFrame: attributeString(child, "office:target-frame-name"),
              children: parseRuns(child, context),
            },
          },
        ];
      }
      if (child.name === "text:note") return parseNote(child, context);
      if (
        child.name === "text:alphabetical-index-mark-start" ||
        child.name === "text:user-index-mark-start"
      ) {
        const name =
          attributeString(child, "text:id") ??
          attributeString(child, "text:key1") ??
          attributeString(child, "text:title") ??
          "";
        if (!name) throw unknownOdtElement(child);
        return [{ bookmarkStart: { id: bookmarkRunId(context, name), name } }];
      }
      if (
        child.name === "text:alphabetical-index-mark-end" ||
        child.name === "text:user-index-mark-end"
      ) {
        const name =
          attributeString(child, "text:id") ??
          attributeString(child, "text:key1") ??
          attributeString(child, "text:title") ??
          "";
        const id = context.bookmarkIds.get(name);
        if (!name || id === undefined) throw unknownOdtElement(child);
        return [{ bookmarkEnd: { id } }];
      }
      if (
        child.name === "text:alphabetical-index-mark" ||
        child.name === "text:user-index-mark" ||
        child.name === "text:reference-mark"
      ) {
        const name =
          attributeString(child, "text:name") ??
          attributeString(child, "text:key1") ??
          attributeString(child, "text:title") ??
          "";
        if (!name) throw unknownOdtElement(child);
        return [{ bookmark: { name } }];
      }
      if (child.name === "text:bookmark") {
        return [{ bookmark: { name: attributeString(child, "text:name") ?? "" } }];
      }
      if (child.name === "text:bookmark-start" || child.name === "text:reference-mark-start") {
        const name = attributeString(child, "text:name") ?? "";
        if (!name || context.bookmarkIds.has(name)) throw unknownOdtElement(child);
        return [{ bookmarkStart: { id: bookmarkRunId(context, name), name } }];
      }
      if (child.name === "text:bookmark-end" || child.name === "text:reference-mark-end") {
        const name = attributeString(child, "text:name") ?? "";
        const id = context.bookmarkIds.get(name);
        if (!name || id === undefined) throw unknownOdtElement(child);
        return [{ bookmarkEnd: { id } }];
      }
      if (child.name === "text:change-start") {
        const id = attributeString(child, "text:change-id") ?? "";
        const endIndex = nodes.findIndex(
          (node, nodeIndex) =>
            nodeIndex > index &&
            node.name === "text:change-end" &&
            attributeString(node, "text:change-id") === id,
        );
        const insertion =
          context.changeAlternatives.get(`${id}:insertion`) ?? context.changes.get(id);
        const format = context.changeAlternatives.get(`${id}:format`) ?? context.changes.get(id);
        const change = insertion?.kind === "insertion" ? insertion : format;
        if (change?.kind === "format") {
          if (endIndex < 0) throw malformedOdtElement(child.name, "format change end is missing");
          const changedElement: Element = {
            ...nodes[endIndex]!,
            elements: nodes.slice(index + 1, endIndex),
          };
          resumeAfter = endIndex;
          return formatChangeChildren(parseRuns(changedElement, context), {
            id: change.id,
            author: change.author,
            date: change.date,
          });
        }
        if (change?.kind !== "insertion") throw unknownOdtElement(child);
        if (endIndex < 0) {
          const insertionBody: Element = { ...child, elements: nodes.slice(index + 1) };
          const parsedChildren = parseRuns(insertionBody, context);
          const first = parsedChildren[0];
          const children = (
            parsedChildren.length === 1 &&
            typeof first === "object" &&
            "text" in first &&
            first.text !== undefined
              ? [first.text]
              : parsedChildren
          ).filter(isTrackChangeChild);
          if (children.length !== parsedChildren.length) throw unknownOdtElement(child);
          context.pendingInsertions.add(id);
          resumeAfter = nodes.length;
          return [
            { insertion: { id: change.id, author: change.author, date: change.date, children } },
          ];
        }
        const insertionBody: Element = {
          ...nodes[endIndex]!,
          elements: nodes.slice(index + 1, endIndex),
        };
        const parsedChildren = parseRuns(insertionBody, context);
        const first = parsedChildren[0];
        const children = (
          parsedChildren.length === 1 &&
          typeof first === "object" &&
          "text" in first &&
          first.text !== undefined
            ? [first.text]
            : parsedChildren
        ).filter(isTrackChangeChild);
        if (children.length !== parsedChildren.length) throw unknownOdtElement(child);
        resumeAfter = endIndex;
        return [
          {
            insertion: {
              id: change.id,
              author: change.author,
              date: change.date,
              children,
            },
          },
        ];
      }
      if (child.name === "text:change-end") {
        const id = attributeString(child, "text:change-id") ?? "";
        context.pendingInsertions.delete(id);
        return [];
      }
      if (child.name === "text:change") {
        const id = attributeString(child, "text:change-id") ?? "";
        const change = context.changeAlternatives.get(`${id}:deletion`) ?? context.changes.get(id);
        if (change?.kind !== "deletion") throw unknownOdtElement(child);
        const parsedChildren = change.children?.length
          ? parseRuns(change.children[0]!, context)
          : [];
        const first = parsedChildren[0];
        const children = (
          parsedChildren.length === 1 &&
          typeof first === "object" &&
          "text" in first &&
          first.text !== undefined
            ? [first.text]
            : parsedChildren
        ).filter(isTrackChangeChild);
        if (children.length !== parsedChildren.length) throw unknownOdtElement(child);
        return [
          {
            deletion: {
              id: change.id,
              author: change.author,
              date: change.date,
              children,
            },
          },
        ];
      }
      if (child.name === "text:format-change") throw unknownOdtElement(child);
      if (child.name === "text:bookmark-ref" || child.name === "text:reference-ref") {
        const format = fieldReferenceSwitch(attributeString(child, "text:reference-format"));
        return [
          { simpleField: { instruction: `REF ${encodeFieldArgument(textOf(child))}${format}` } },
        ];
      }
      if (child.name === "text:chapter") return [chapterField(parseChapter(child))];
      if (child.name === "text:expression") return [expressionField(parseExpression(child))];
      if (child.name === "text:variable-get") {
        return [
          {
            simpleField: {
              instruction: `ODF VARIABLEGET ${encodeFieldArgument(attributeString(child, "text:name") ?? "")}`,
              cachedValue: textOf(child) || undefined,
            },
          },
        ];
      }
      if (child.name === "text:user-field-get") {
        return [
          {
            simpleField: {
              instruction: `ODF USERFIELD ${encodeFieldArgument(attributeString(child, "text:name") ?? "")}`,
              cachedValue: textOf(child) || undefined,
            },
          },
        ];
      }
      if (child.name === "text:user-field-input") {
        return [
          {
            simpleField: {
              instruction: `ODF USERFIELDINPUT ${encodeFieldArgument(attributeString(child, "text:name") ?? "")}`,
              cachedValue: textOf(child) || undefined,
            },
          },
        ];
      }
      if (child.name === "text:variable-set") {
        const value = parseVariableSet(child);
        const expected = context.pendingVariables.get(value.name);
        if (expected !== undefined && expected !== value.valueType) throw unknownOdtElement(child);
        context.pendingVariables.delete(value.name);
        return [variableField(value)];
      }
      if (child.name === "text:sequence") {
        const value = parseSequence(child);
        context.pendingSequences.delete(value.name);
        return [sequenceField(value)];
      }
      if (child.name === "text:sequence-ref") {
        const referenceName = attributeString(child, "text:ref-name") ?? "";
        return [
          {
            simpleField: {
              instruction: `ODF SEQUENCEREF ${encodeFieldArgument(referenceName)}`,
              cachedValue: textOf(child) || undefined,
            },
          },
        ];
      }
      if (child.name === "text:page-number") {
        const select = attributeString(child, "text:select-page");
        const adjust = attributeNumber(child, "text:page-adjust");
        return [
          {
            simpleField: {
              instruction: [
                "PAGE",
                select ? `\\select ${select}` : undefined,
                adjust === undefined ? undefined : `\\adjust ${adjust}`,
              ]
                .filter(Boolean)
                .join(" "),
              cachedValue: textOf(child) || undefined,
            },
          },
        ];
      }
      if (child.name === "text:page-count") {
        return [
          { simpleField: { instruction: "NUMPAGES", cachedValue: textOf(child) || undefined } },
        ];
      }
      if (child.name === "text:date") {
        const style = attributeString(child, "style:data-style-name");
        return [
          {
            simpleField: {
              instruction: `DATE${style ? ` \\@ ${encodeFieldArgument(style)}` : ""}`,
              cachedValue: textOf(child) || undefined,
            },
          },
        ];
      }
      if (child.name === "text:author-name") {
        return [
          { simpleField: { instruction: "AUTHOR", cachedValue: textOf(child) || undefined } },
        ];
      }
      const senderSuffix = SENDER_FIELD_SUFFIXES[child.name ?? ""];
      if (senderSuffix) {
        return [
          {
            simpleField: {
              instruction: `AUTHOR \\${senderSuffix}`,
              cachedValue: textOf(child) || undefined,
            },
          },
        ];
      }
      if (child.name === "text:title") {
        return [
          {
            simpleField: {
              instruction: "DOCPROPERTY Title",
              cachedValue: textOf(child) || undefined,
            },
          },
        ];
      }
      if (child.name === "text:user-defined") {
        const name = attributeString(child, "text:name") ?? "";
        return [
          {
            simpleField: {
              instruction: `DOCPROPERTY ${encodeFieldArgument(name)}`,
              cachedValue: textOf(child) || undefined,
            },
          },
        ];
      }
      if (child.name === "text:database-display") {
        const name = attributeString(child, "text:column-name") ?? "";
        return [
          {
            simpleField: {
              instruction: `MERGEFIELD ${encodeFieldArgument(name)}`,
              cachedValue: textOf(child) || undefined,
            },
          },
        ];
      }
      if (child.name === "text:placeholder") {
        const type = attributeString(child, "text:placeholder-type") ?? "text";
        const description = attributeString(child, "text:description");
        return [
          {
            simpleField: {
              instruction: `ODF PLACEHOLDER ${encodeFieldArgument(type)}${description ? ` ${encodeFieldArgument(description)}` : ""}`,
              cachedValue: textOf(child) || undefined,
            },
          },
        ];
      }
      if (child.name === "text:text-input") {
        return [
          { simpleField: { instruction: "FILLIN", cachedValue: textOf(child) || undefined } },
        ];
      }
      if (child.name === "text:hidden-paragraph") {
        const condition = attributeString(child, "text:condition") ?? "";
        return [
          { simpleField: { instruction: `ODF HIDDENPARAGRAPH ${encodeFieldArgument(condition)}` } },
        ];
      }
      if (child.name === "text:drop-down") {
        const entries = childrenNamed(child, "text:label").map((label) => textOf(label));
        return [
          {
            formField: {
              name: attributeString(child, "text:name") ?? "",
              dropDownList: { entries },
            },
          },
        ];
      }
      if (child.name === "loext:content-control") {
        const entries = childrenNamed(child, "loext:list-item").map((item) => textOf(item));
        return [
          {
            formField: {
              name: attributeString(child, "loext:id") ?? "",
              ...(entries.length
                ? { dropDownList: { entries } }
                : { textInput: { value: textOf(child) || undefined } }),
            },
          },
        ];
      }
      if (child.name === "loext:style-ref") {
        const styleName = attributeString(child, "text:ref-name") ?? "";
        const format = attributeString(child, "text:reference-format") ?? "text";
        return [
          {
            simpleField: {
              instruction: `ODF STYLEREF ${encodeFieldArgument(format)} ${encodeFieldArgument(styleName)}`,
              cachedValue: textOf(child) || undefined,
            },
          },
        ];
      }
      if (child.name === "loext:page-count-range") {
        const format = attributeString(child, "style:num-format") ?? "";
        return [
          {
            simpleField: {
              instruction: `SECTIONPAGES${format ? ` \\@ ${encodeFieldArgument(format)}` : ""}`,
              cachedValue: textOf(child) || undefined,
            },
          },
        ];
      }
      if (child.name === "text:meta" || child.name === "text:meta-field") {
        const name = attributeString(child, "text:name") ?? "";
        return [
          {
            simpleField: {
              instruction: `ODF META ${encodeFieldArgument(name)}`,
              cachedValue: textOf(child) || undefined,
            },
          },
        ];
      }
      if (child.name === "text:hidden-text") return [hiddenTextField(parseHiddenText(child))];
      if (child.name === "text:bibliography-mark")
        return [
          {
            simpleField: {
              instruction: `CITATION ${encodeFieldArgument(
                attributeString(child, "text:bibliography-type") ?? "custom",
              )}`,
              cachedValue: textOf(child) || undefined,
            },
          },
        ];
      if (child.name === "office:annotation") return parseAnnotationStart(child, context);
      if (child.name === "office:annotation-end") return parseAnnotationEnd(child, context);
      if (child.name === "draw:a") return [parseShapeHyperlink(child)];
      if (isShapeElement(child.name)) return [{ wpsShape: parseShape(child) }];
      if (child.name === "text:span") {
        const styleEntry = context.styles.get(attributeString(child, "text:style-name") ?? "");
        const properties = styleEntry?.character;
        return [
          {
            text: textOf(child),
            ...(styleEntry?.properties?.some(
              (property) => Object.keys(property.attributes).length > 0,
            )
              ? { style: attributeString(child, "text:style-name") }
              : {}),
            ...properties,
            underline: properties?.underline ? { type: "single" } : undefined,
          },
        ];
      }
      if (child.name === "text:ruby") return [{ text: "", children: [parseRuby(child, context)] }];
      if (child.name === "text:tab") return [{ text: "", children: [{ tab: true }] }];
      if (child.name === "draw:frame") {
        if (childNamed(child, "draw:object-ole")) return [];
        if (childNamed(child, "draw:text-box")) return [parseTextBoxFrame(child, context)];
        return childNamed(child, "draw:object")
          ? parseChartFrame(child, context)
          : parsePictureFrame(child, context);
      }
      if (child.name === "draw:object") return parseChartFrame(child, context);
      if (child.name === "draw:g") return [parseGroup(child, context)];
      if (child.name === "draw:custom-shape") {
        const shape = parseCustomShape(child, context);
        return shape ? [shape] : [];
      }
      if (child.name) throw unknownOdtElement(child);
      return [];
    },
  );
  return runs.reduce<(string | RunOptions | ParagraphChild)[]>((merged, run) => {
    const last = merged.at(-1);
    if (typeof last === "string" && typeof run === "string") merged[merged.length - 1] = last + run;
    else merged.push(run);
    return merged;
  }, []);
}

function runStyleName(children: CanonicalRuby["base"]["children"]): string | undefined {
  const first = children?.find(
    (child): child is RunOptions => typeof child === "object" && "text" in child,
  );
  return first?.style;
}

function rubyContentText(content: CanonicalRuby["base"]): string {
  return (content.children ?? [])
    .map((child) => {
      if (typeof child === "string") return child;
      if ("text" in child && typeof child.text === "string") return child.text;
      return "";
    })
    .join("");
}

function rubyContentXml(
  content: CanonicalRuby["base"],
  styles: string[],
  images: OdtImage[],
  notes: NotesContext,
  numbering: DocumentOptions["numbering"],
  charts: OdtChart[],
  embeddedObjects: OdtEmbeddedObjectOptions[],
): string {
  return runXml(
    { children: content.children },
    styles,
    images,
    notes,
    numbering,
    charts,
    embeddedObjects,
  ).join("");
}

function rubyXml(
  ruby: CanonicalRuby,
  styles: string[],
  images: OdtImage[],
  notes: NotesContext,
  numbering: DocumentOptions["numbering"],
  charts: OdtChart[],
  embeddedObjects: OdtEmbeddedObjectOptions[],
): string {
  return xmlElement(
    "text:ruby",
    { "text:style-name": ruby.style ?? runStyleName(ruby.base.children) },
    [
      xmlElement("text:ruby-base", undefined, [
        rubyContentXml(ruby.base, styles, images, notes, numbering, charts, embeddedObjects),
      ]),
      xmlElement("text:ruby-text", { "text:style-name": runStyleName(ruby.text.children) }, [
        escapeText(rubyContentText(ruby.text)),
      ]),
    ],
  );
}

function parseRuby(element: Element, context: ParseContext): RunChild {
  const baseElement = childNamed(element, "text:ruby-base");
  const textElement = childNamed(element, "text:ruby-text");
  if (!baseElement || !textElement) throw unknownOdtElement(element);
  const styleEntry = context.styles.get(attributeString(element, "text:style-name") ?? "");
  const rubyProperties = styleEntry?.properties?.find(
    (property) => property.name === "style:ruby-properties",
  );
  const alignmentAttribute = rubyProperties?.attributes["style:ruby-align"];
  const alignments = {
    left: "left",
    center: "center",
    right: "right",
    "distribute-letter": "distributeLetter",
    "distribute-space": "distributeSpace",
  } as const;
  const alignment =
    alignmentAttribute === undefined
      ? "center"
      : alignments[String(alignmentAttribute) as keyof typeof alignments];
  if (!alignment) throw unknownOdtElement(element);
  const annotationStyle = context.styles.get(attributeString(textElement, "text:style-name") ?? "");
  const baseChildren = parseRubyContent(baseElement, context);
  const rubyStyle = attributeString(element, "text:style-name");
  const firstBaseChild = baseChildren.find(
    (child): child is RunOptions => typeof child === "object" && "text" in child,
  );
  if (rubyStyle && firstBaseChild && firstBaseChild.style === undefined)
    firstBaseChild.style = rubyStyle;
  const textProperties = annotationStyle?.properties?.find(
    (property) => property.name === "style:text-properties",
  )?.attributes;
  const language = textProperties?.["fo:language"];
  const country = textProperties?.["fo:country"];
  const languageId =
    typeof language === "string" && typeof country === "string"
      ? `${language}-${country}`
      : typeof language === "string"
        ? language
        : "en-US";
  return {
    ruby: {
      properties: {
        alignment,
        fontSize: annotationStyle?.character.size ?? 10,
        raise: 0,
        baseFontSize: firstBaseChild
          ? (context.styles.get(firstBaseChild.style ?? "")?.character.size ?? 10)
          : 10,
        languageId,
      },
      style: attributeString(element, "text:style-name"),
      text: {
        children: [
          { text: textOf(textElement), style: attributeString(textElement, "text:style-name") },
        ],
      },
      base: { children: baseChildren },
    },
  };
}

function parseRubyContent(
  element: Element,
  context: ParseContext,
): NonNullable<CanonicalRuby["base"]["children"]> {
  const children: NonNullable<CanonicalRuby["base"]["children"]> = [];
  for (const child of parseRuns(element, context)) {
    if (typeof child === "string") {
      children.push(child);
      continue;
    }
    children.push(child);
  }
  return children;
}

function parseFieldmark(element: Element, value: string, context: ParseContext): ParagraphChild {
  const type = attributeString(element, "field:type") ?? "";
  const parameters = childrenNamed(element, "field:param");
  const parameter = (name: string): string | undefined =>
    parameters
      .find((parameterElement) => attributeString(parameterElement, "field:name") === name)
      ?.attributes?.["field:value"]?.toString();
  const name =
    parameter("Name") ??
    attributeString(element, "text:name")?.replace(/^__Fieldmark__\d+_/, "") ??
    "";
  if (type.endsWith("FORMCHECKBOX")) {
    const field = {
      formField: {
        name,
        checkBox: { checked: parameter("Checkbox_Checked") === "true" },
      },
    };
    context.notes.forms.containers.push({
      controls: [
        {
          control: "checkbox",
          name,
          currentState: parameter("Checkbox_Checked") === "true" ? "checked" : "unchecked",
        },
      ],
    });
    return field;
  }
  if (type.endsWith("FORMDROPDOWN")) {
    const entries = parameters
      .filter(
        (parameterElement) =>
          attributeString(parameterElement, "field:name") === "Dropdown_ListEntry",
      )
      .map((parameterElement) => String(parameterElement.attributes?.["field:value"] ?? ""));
    const selected = Number(parameter("Dropdown_Selected"));
    const selectedIndex = Number.isInteger(selected) && selected >= 0 ? selected : undefined;
    context.notes.forms.containers.push({
      controls: [
        {
          control: "listBox",
          name,
          options: entries.map((label, index) => ({
            label,
            currentSelected: index === selectedIndex,
          })),
        },
      ],
    });
    return {
      formField: {
        name,
        dropDownList: {
          entries,
          ...(selectedIndex !== undefined ? { result: selectedIndex } : {}),
        },
      },
    };
  }
  if (
    type.endsWith("FORMTEXT") ||
    type.endsWith("FORMDATE") ||
    type.endsWith("FORMTIME") ||
    type.endsWith("FORMNUMBER") ||
    type === ""
  ) {
    const format =
      parameter("DateField_DateFormat") ??
      parameter("TimeField_TimeFormat") ??
      parameter("NumberField_NumberFormat");
    context.notes.forms.containers.push({
      controls: [
        {
          control: "text",
          name,
          ...(value !== undefined ? { currentValue: value } : {}),
        },
      ],
    });
    return {
      formField: {
        name,
        textInput: {
          ...(value ? { value } : {}),
          ...(format ? { format } : {}),
          ...(parameter("MaxLength") ? { maxLength: Number(parameter("MaxLength")) } : {}),
        },
      },
    };
  }
  throw unknownOdtElement(element);
}

export function bookmarkRunId(context: ParseContext, name: string): number {
  const existing = context.bookmarkIds.get(name);
  if (existing !== undefined) return existing;
  const id = nextBookmarkId(context);
  context.bookmarkIds.set(name, id);
  return id;
}

export function encodeFieldArgument(value: string): string {
  return value.replace(/([\\"]|\s)/g, "\\$1");
}

export function variableField(value: ReturnType<typeof parseVariableSet>): {
  simpleField: CanonicalSimpleField;
} {
  const formula = value.formula ? ` = ${value.formula}` : "";
  return {
    simpleField: {
      instruction: `VARIABLE ${encodeFieldArgument(value.name)}${formula}`,
      cachedValue: value.display ?? value.value,
    },
  };
}

export function sequenceField(value: ReturnType<typeof parseSequence>): {
  simpleField: CanonicalSimpleField;
} {
  return {
    simpleField: {
      instruction: `SEQ ${encodeFieldArgument(value.name)}${value.formula ? ` = ${value.formula}` : ""}`,
      cachedValue: value.display,
    },
  };
}

export function chapterField(value: ReturnType<typeof parseChapter>): {
  simpleField: CanonicalSimpleField;
} {
  return {
    simpleField: {
      instruction: `STYLEREF ${value.outlineLevel ?? 1} \\n`,
      cachedValue: value.display,
    },
  };
}

export function expressionField(value: ReturnType<typeof parseExpression>): {
  simpleField: CanonicalSimpleField;
} {
  return {
    simpleField: {
      instruction: `= ${value.formula ?? ""}`.trim(),
      cachedValue: value.display ?? value.value,
    },
  };
}

export function hiddenTextField(value: ReturnType<typeof parseHiddenText>): {
  simpleField: CanonicalSimpleField;
} {
  if (value.fixed) {
    throw malformedOdtElement("text:hidden-text", "fixed hidden text has no canonical mapping");
  }
  return {
    simpleField: {
      instruction: `IF ${encodeFieldArgument(value.condition)} "${value.content.replace(
        /([\\"])/g,
        "\\$1",
      )}" ""`,
      cachedValue: value.hidden ? undefined : value.content,
    },
  };
}

export function parseVariableSet(element: Element): {
  name: string;
  valueType: "float" | "string";
  value?: string;
  display?: string;
  formula?: string;
} {
  const rawType = attributeString(element, "office:value-type") ?? "string";
  if (rawType !== "float" && rawType !== "string") throw unknownOdtElement(element);
  return {
    name: attributeString(element, "text:name") ?? "",
    valueType: rawType,
    value: attributeString(element, "office:value"),
    display: textOf(element) || attributeString(element, "text:display"),
    formula: attributeString(element, "text:formula"),
  };
}

export function parseSequence(element: Element): {
  name: string;
  referenceName?: string;
  display?: string;
  formula?: string;
} {
  return {
    name: attributeString(element, "text:name") ?? "",
    referenceName: attributeString(element, "text:ref-name"),
    display: textOf(element) || attributeString(element, "text:display"),
    formula: attributeString(element, "text:formula"),
  };
}

export function parseChapter(element: Element): { display?: string; outlineLevel?: number } {
  return {
    display: textOf(element) || attributeString(element, "text:display"),
    outlineLevel: attributeNumber(element, "text:outline-level"),
  };
}

export function parseExpression(element: Element): {
  formula?: string;
  valueType: "float" | "string";
  value?: string;
  display?: string;
} {
  const rawType = attributeString(element, "office:value-type") ?? "string";
  if (rawType !== "float" && rawType !== "string") throw unknownOdtElement(element);
  return {
    formula: attributeString(element, "text:formula"),
    valueType: rawType,
    value: attributeString(element, "office:value"),
    display: textOf(element) || attributeString(element, "text:display"),
  };
}

export function parseHiddenText(element: Element): {
  condition: string;
  content: string;
  hidden?: boolean;
  fixed?: boolean;
} {
  return {
    condition: attributeString(element, "text:condition") ?? "",
    content: textOf(element),
    hidden: attributeString(element, "text:is-hidden") === "true" ? true : undefined,
    fixed: attributeString(element, "text:is-fixed") === "true" ? true : undefined,
  };
}

export function parseSequenceDeclarations(element: Element, context: ParseContext): void {
  assertElementNames(element, [], ["text:sequence-decl"]);
  for (const declaration of childrenNamed(element, "text:sequence-decl")) {
    assertElementNames(declaration, [
      "text:name",
      "text:display-outline-level",
      "text:outline-level-separator",
      "text:separation-character",
    ]);
    const name = attributeString(declaration, "text:name");
    if (!name) throw unknownOdtElement(declaration);
    const level = attributeNumber(declaration, "text:display-outline-level") ?? 0;
    const outlineSeparator =
      attributeString(declaration, "text:outline-level-separator") ??
      attributeString(declaration, "text:separation-character");
    context.declaredSequences.push({
      name,
      ...(level ? { displayOutlineLevel: level } : {}),
      ...(outlineSeparator !== undefined ? { outlineSeparator } : {}),
    });
    context.pendingSequences.add(name);
  }
}

export function parseVariableDeclarations(element: Element, context: ParseContext): void {
  assertElementNames(element, [], ["text:variable-decl"]);
  for (const declaration of childrenNamed(element, "text:variable-decl")) {
    assertElementNames(declaration, ["text:name", "office:value-type"]);
    const name = attributeString(declaration, "text:name");
    const valueType = attributeString(declaration, "office:value-type") ?? "string";
    if (!name || (valueType !== "float" && valueType !== "string")) {
      throw unknownOdtElement(declaration);
    }
    context.declaredVariables.push({ name, valueType });
    context.pendingVariables.set(name, valueType);
  }
}

export function parseUserFieldDeclarations(element: Element, context: ParseContext): void {
  assertElementNames(element, [], ["text:user-field-decl"]);
  for (const declaration of childrenNamed(element, "text:user-field-decl")) {
    assertElementNames(declaration, [
      "text:name",
      "text:formula",
      "office:value-type",
      "office:value",
      "office:boolean-value",
      "office:date-value",
      "office:string-value",
      "office:time-value",
    ]);
    const name = attributeString(declaration, "text:name");
    if (!name) throw unknownOdtElement(declaration);
    context.documentVariables.push({
      name,
      val:
        attributeString(declaration, "office:string-value") ??
        attributeString(declaration, "office:value") ??
        attributeString(declaration, "office:boolean-value") ??
        attributeString(declaration, "office:date-value") ??
        attributeString(declaration, "office:time-value") ??
        "",
    });
  }
}

export function parseDdeConnectionDeclarations(element: Element, context: ParseContext): void {
  assertElementNames(element, [], ["text:dde-connection-decl"]);
  for (const declaration of childrenNamed(element, "text:dde-connection-decl")) {
    assertElementNames(declaration, [
      "office:name",
      "office:dde-application",
      "office:dde-topic",
      "office:dde-item",
      "office:automatic-update",
    ]);
    const name = attributeString(declaration, "office:name");
    const application = attributeString(declaration, "office:dde-application") ?? "";
    const topic = attributeString(declaration, "office:dde-topic") ?? "";
    const item = attributeString(declaration, "office:dde-item") ?? "";
    if (!name || !application || !topic || !item) throw unknownOdtElement(declaration);
    context.documentVariables.push({ name: `DDE ${name}`, val: `${application}|${topic}|${item}` });
  }
}

export function parseTrackedChanges(element: Element, context: ParseContext): void {
  assertElementNames(element, ["text:track-changes"], ["text:changed-region"]);
  context.trackRevisions = attributeString(element, "text:track-changes") !== "false";
  for (const region of childrenNamed(element, "text:changed-region")) {
    assertElementNames(
      region,
      ["xml:id", "text:id"],
      ["text:insertion", "text:deletion", "text:format-change"],
    );
    const id = attributeString(region, "text:id") ?? attributeString(region, "xml:id") ?? "";
    const children =
      region.elements?.filter((child) =>
        ["text:insertion", "text:deletion", "text:format-change"].includes(child.name ?? ""),
      ) ?? [];
    if (children.length < 1) throw unknownOdtElement(region);
    const kinds = new Set<string>();
    let hasPrimary = false;
    for (const change of children) {
      const kind =
        change.name === "text:deletion"
          ? "deletion"
          : change.name === "text:format-change"
            ? "format"
            : "insertion";
      if (kinds.has(kind)) throw unknownOdtElement(region);
      kinds.add(kind);
      const info = childNamed(change, "office:change-info");
      const author = info ? textOf(childNamed(info, "dc:creator")) : "";
      const date = info ? textOf(childNamed(info, "dc:date")) : "";
      if (!author || !date) throw unknownOdtElement(change);
      const content =
        kind === "deletion"
          ? (change.elements ?? []).filter((child) => child.name !== "office:change-info")
          : [];
      const metadata = (change.elements ?? []).filter(
        (child) => child.name === "office:change-info",
      );
      if (metadata.length !== 1 || (kind !== "deletion" && content.length !== 0))
        throw unknownOdtElement(change);
      const deletedParagraphs = content.flatMap((child) =>
        child.name === "text:p"
          ? [child]
          : childrenNamed(child, "text:p").flatMap((listItem) => childrenNamed(listItem, "text:p")),
      );
      const parsed: ParsedTrackChange = {
        id: nextTrackChangeId(context),
        author,
        date,
        kind,
        children: deletedParagraphs,
      };
      if (hasPrimary) context.changeAlternatives.set(`${id}:${kind}`, parsed);
      else context.changes.set(id, parsed);
      hasPrimary = true;
    }
  }
}
