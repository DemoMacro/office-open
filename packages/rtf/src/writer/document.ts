import type {
  DocumentOptions,
  ParagraphOptions,
  RunOptions,
  SectionChild,
  SectionOptions,
} from "@office-open/docx";

import { assertAllowed, createContext, reject, supportedMetadata } from "./context";
import { control, rtfText } from "./escape";
import { writeColorTable, writeFontTable } from "./font-color";
import { writeNumberingTable, type NumberingReferences } from "./numbering";
import { writeParagraph } from "./paragraph";
import { writeStyles } from "./styles";
import { writeTable } from "./table";

function destinationText(paragraph: ParagraphOptions, path: string): string {
  assertAllowed(paragraph, ["children"], "header-footer", path);
  return (paragraph.children ?? [])
    .map((run, index) => {
      assertAllowed(run, ["text"], "header-footer", `${path}.children[${index}]`);
      return rtfText((run as RunOptions).text ?? "");
    })
    .join("");
}

function writeDestinations(section: SectionOptions, path: string): string {
  let result = "";
  for (const kind of ["headers", "footers"] as const) {
    const destinations = section[kind];
    if (!destinations) continue;
    assertAllowed(destinations, ["default", "first"], kind, `${path}.${kind}`);
    for (const slot of ["default", "first"] as const) {
      const blocks = destinations[slot];
      if (blocks === undefined) continue;
      const suffix = slot === "first" ? "f" : "";
      result += `{\\${kind.slice(0, -1)}${suffix} `;
      for (const [index, child] of blocks.entries()) {
        if (!("paragraph" in child) || typeof child.paragraph === "string") {
          reject(
            kind,
            `${path}.${kind}.${slot}[${index}]`,
            "child",
            "RTF destination projections contain paragraphs",
          );
        }
        result += destinationText(child.paragraph, `${path}.${kind}.${slot}[${index}].paragraph`);
        result += control("par");
      }
      result += "}";
    }
  }
  return result;
}

function writeFootnotes(options: DocumentOptions): string {
  if (!options.footnotes) return "";
  return options.footnotes
    .map((footnote, index) => {
      assertAllowed(footnote, ["children"], "footnote", `footnotes[${index}]`);
      const blocks = footnote.children
        .map((child, childIndex) => {
          if (
            typeof child !== "object" ||
            child === null ||
            !("paragraph" in child) ||
            typeof child.paragraph === "string"
          ) {
            reject(
              "footnote",
              `footnotes[${index}].children[${childIndex}]`,
              "child",
              "RTF footnotes contain paragraphs",
            );
          }
          return `${destinationText(
            child.paragraph,
            `footnotes[${index}].children[${childIndex}].paragraph`,
          )}${control("par")}`;
        })
        .join("");
      return `{\\footnote ${blocks}}`;
    })
    .join("");
}

function writeMetadata(options: DocumentOptions): string {
  const metadata = supportedMetadata(options);
  const entries = Object.entries(metadata).map(([name, value]) => {
    const word = name === "creator" ? "author" : name === "description" ? "doccomm" : name;
    return `{\\${word} ${rtfText(value)}}`;
  });
  return entries.length > 0 ? `{\\info${entries.join("")}}` : "";
}

function writeSectionProperties(section: SectionOptions, path: string): string {
  const properties = section.properties;
  if (!properties) return "";
  assertAllowed(properties, ["pageSize", "pageMargin"], "section", `${path}.properties`);
  let result = "";
  if (properties.pageSize !== false && properties.pageSize) {
    assertAllowed(
      properties.pageSize,
      ["width", "height"],
      "section",
      `${path}.properties.pageSize`,
    );
    const width = properties.pageSize.width;
    const height = properties.pageSize.height;
    if (width !== undefined && typeof width !== "number") {
      reject(
        "section",
        `${path}.properties.pageSize`,
        "width",
        "RTF page dimensions require numeric twips",
      );
    }
    if (height !== undefined && typeof height !== "number") {
      reject(
        "section",
        `${path}.properties.pageSize`,
        "height",
        "RTF page dimensions require numeric twips",
      );
    }
    result += width === undefined ? "" : control("pgwsxn", width);
    result += height === undefined ? "" : control("pghsxn", height);
  }
  if (properties.pageMargin !== false && properties.pageMargin) {
    assertAllowed(
      properties.pageMargin,
      ["top", "bottom", "left", "right"],
      "section",
      `${path}.properties.pageMargin`,
    );
    const margin = properties.pageMargin;
    for (const [field, word] of [
      ["left", "marglsxn"],
      ["right", "margsxn"],
      ["top", "margtsxn"],
      ["bottom", "margbsxn"],
    ] as const) {
      const value = margin[field];
      if (value === undefined) continue;
      if (typeof value !== "number") {
        reject(
          "section",
          `${path}.properties.pageMargin`,
          field,
          "RTF margins require numeric twips",
        );
      }
      result += control(word, value);
    }
  }
  return result;
}

function writeChild(
  child: SectionChild,
  context: ReturnType<typeof createContext>,
  references: NumberingReferences,
  path: string,
): string {
  if ("paragraph" in child) {
    if (typeof child.paragraph === "string") {
      assertAllowed(child, ["paragraph"], "document", path);
      return `${rtfText(child.paragraph)}${control("par")}`;
    }
    return writeParagraph(child.paragraph, context, references, `${path}.paragraph`);
  }
  if ("table" in child) return writeTable(child.table, context, references, `${path}.table`);
  reject(
    "document",
    path,
    Object.keys(child)[0] ?? "child",
    "RTF writer does not support this section child",
  );
}

function writeSections(
  options: DocumentOptions,
  references: NumberingReferences,
  context: ReturnType<typeof createContext>,
): string {
  return options.sections
    .map((section, index) => {
      const path = `sections[${index}]`;
      assertAllowed(section, ["headers", "footers", "properties", "children"], "section", path);
      const children = section.children;
      const content = children
        .map((child, childIndex) =>
          writeChild(child, context, references, `${path}.children[${childIndex}]`),
        )
        .join("");
      const boundary = control("par");
      const body =
        index === options.sections.length - 1 || !content.endsWith(boundary)
          ? content
          : content.slice(0, -boundary.length);
      const separator = index === options.sections.length - 1 ? "" : control("sect");
      return `${writeSectionProperties(section, path)}${body}${separator}`;
    })
    .join("");
}

export function writeRtf(options: DocumentOptions): string {
  assertAllowed(
    options,
    [
      "sections",
      "title",
      "subject",
      "creator",
      "keywords",
      "description",
      "category",
      "styles",
      "numbering",
      "footnotes",
    ],
    "document",
    "options",
  );
  const references: NumberingReferences = new Map();
  const context = createContext();
  const first = options.sections[0] ?? { children: [] };
  const body = [
    writeMetadata(options),
    writeStyles(options.styles, context),
    writeNumberingTable(options.numbering, references),
    writeDestinations(first, "sections[0]"),
    writeFootnotes(options),
    writeSections(options, references, context),
  ].join("");
  const tables = `${writeFontTable(context)}${writeColorTable(context)}`;
  return `{\\rtf1${tables}${body}}`;
}
