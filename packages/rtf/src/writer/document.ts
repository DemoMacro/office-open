import { convertToTwip } from "@office-open/core";
import type {
  DocumentOptions,
  ParagraphOptions,
  RunOptions,
  SectionChild,
  SectionOptions,
} from "@office-open/docx";

import { tagToLcid } from "../langs";
import { assertAllowed, createContext, registerFont, reject, supportedMetadata } from "./context";
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
    assertAllowed(destinations, ["default", "even", "first"], kind, `${path}.${kind}`);
    for (const slot of ["default", "even", "first"] as const) {
      const blocks = destinations[slot];
      if (blocks === undefined) continue;
      const suffix = slot === "first" ? "f" : slot === "even" ? "l" : "";
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

function writeMetadata(options: DocumentOptions): string {
  if (options.appProperties) {
    assertAllowed(
      options.appProperties,
      ["totalTime", "pages", "words", "characters", "charactersWithSpaces"],
      "app-properties",
      "appProperties",
    );
  }
  const metadata = supportedMetadata(options);
  const entries = Object.entries(metadata).map(([name, value]) => {
    const word =
      name === "creator"
        ? "author"
        : name === "lastModifiedBy"
          ? "operator"
          : name === "description"
            ? "doccomm"
            : name;
    return `{\\${word} ${rtfText(value)}}`;
  });
  const appProperties = options.appProperties ?? {};
  const appCounts: [keyof typeof appProperties, string][] = [
    ["totalTime", "edmins"],
    ["pages", "nofpages"],
    ["words", "nofwords"],
    ["characters", "nofchars"],
    ["charactersWithSpaces", "nofcharsws"],
  ];
  const appEntries = appCounts
    .map(([field, word]) => {
      const key = String(field);
      const value = appProperties[field];
      if (value === undefined) return "";
      if (typeof value !== "number" || !Number.isInteger(value) || value < 0) {
        reject(
          "app-properties",
          `appProperties.${key}`,
          typeof value === "symbol" ? String(value) : (JSON.stringify(value) ?? ""),
          "RTF requires a count",
        );
        return "";
      }
      return `{\\${word}${value}}`;
    })
    .join("");
  return entries.length > 0 || appEntries ? `{\\info${entries.join("")}${appEntries}}` : "";
}

function writeDefaultLanguage(options: DocumentOptions): string {
  if (options.language === undefined) return "";
  const lcid = tagToLcid(options.language);
  if (lcid === undefined) {
    reject("metadata", "language", options.language, "RTF language table has no such tag");
  }
  return control("deflang", lcid);
}

function writeApplication(options: DocumentOptions): string {
  return options.application === undefined ? "" : `{\\generator ${rtfText(options.application)}}`;
}

const W3CDTF = /^\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})$/u;

function assertTimestampType(value: string | false | undefined, path: string): void {
  if (value === undefined || value === false || value === "dcterms:W3CDTF") return;
  reject("metadata", path, "type", "RTF timestamps require W3CDTF");
}

function writeTimestamp(word: string, value: unknown, path: string): string {
  const date = value instanceof Date ? value : new Date(typeof value === "string" ? value : "");
  if (
    (typeof value !== "string" && !(value instanceof Date)) ||
    (typeof value === "string" && !W3CDTF.test(value)) ||
    Number.isNaN(date.getTime()) ||
    date.getUTCMilliseconds() !== 0
  ) {
    reject("metadata", path, word, "RTF requires a valid W3CDTF date or Date without milliseconds");
  }
  const pad = (component: number) => component.toString().padStart(2, "0");
  return `{\\${word}\\yr${date.getUTCFullYear()}\\mo${pad(date.getUTCMonth() + 1)}\\dy${pad(
    date.getUTCDate(),
  )}\\hr${pad(date.getUTCHours())}\\min${pad(date.getUTCMinutes())}\\sec${pad(
    date.getUTCSeconds(),
  )}}`;
}

const VIEW_KINDS = {
  none: 0,
  print: 1,
  outline: 2,
  masterPages: 3,
  normal: 4,
  web: 5,
} as const;

function settingEnabled(value: unknown): boolean {
  return value === true || value === "true" || value === "1" || value === "on" || value === "t";
}

function settingDisabled(value: unknown): boolean {
  return value === false || value === "false" || value === "0" || value === "off" || value === "f";
}

function writeSettings(settings: DocumentOptions["settings"], path: string): string {
  if (!settings) return "";
  assertAllowed(
    settings,
    ["defaultTabStop", "view", "zoom", "trackRevisions", "doNotTrackMoves"],
    "settings",
    path,
  );
  let result = "";
  if (settings.defaultTabStop !== undefined) {
    const value = convertToTwip(settings.defaultTabStop);
    if (!Number.isInteger(value) || value <= 0) {
      reject("settings", path, "defaultTabStop", "RTF requires a positive integer tab width");
    }
    result += control("deftab", value);
  }
  if (settings.view !== undefined) result += control("viewkind", VIEW_KINDS[settings.view]);
  if (settings.zoom !== undefined) {
    assertAllowed(settings.zoom, ["percent"], "settings", `${path}.zoom`);
    if (settings.zoom.percent !== undefined) {
      const percent =
        typeof settings.zoom.percent === "number"
          ? settings.zoom.percent
          : Number(settings.zoom.percent.slice(0, -1));
      if (!Number.isInteger(percent) || percent <= 0) {
        reject("settings", `${path}.zoom`, "percent", "RTF requires a positive integer percentage");
      }
      result += control("viewscale", percent);
    }
  }
  if (settings.trackRevisions !== undefined) {
    if (!settingEnabled(settings.trackRevisions) && !settingDisabled(settings.trackRevisions)) {
      reject("settings", path, "trackRevisions", "RTF requires a valid on/off value");
    }
    if (settingEnabled(settings.trackRevisions)) result += control("revisions");
  }
  if (settings.doNotTrackMoves !== undefined) {
    if (!settingEnabled(settings.doNotTrackMoves) && !settingDisabled(settings.doNotTrackMoves)) {
      reject("settings", path, "doNotTrackMoves", "RTF requires a valid on/off value");
    }
    if (settingDisabled(settings.doNotTrackMoves)) result += control("trackmoves");
  }
  return result;
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
  if ("bookmarkStart" in child) {
    assertAllowed(child.bookmarkStart, ["id", "name"], "bookmark", `${path}.bookmarkStart`);
    return `{\\*\\bkmkstart ${rtfText(child.bookmarkStart.name)}}`;
  }
  if ("bookmarkEnd" in child) {
    assertAllowed(child.bookmarkEnd, ["id"], "bookmark", `${path}.bookmarkEnd`);
    return "{\\*\\bkmkend }";
  }
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
      "lastModifiedBy",
      "keywords",
      "description",
      "category",
      "created",
      "createdType",
      "modified",
      "modifiedType",
      "lastPrinted",
      "language",
      "endnotes",
      "settings",
      "rawParts",
      "passthroughRelationships",
      "appProperties",
      "styles",
      "numbering",
      "footnotes",
      "defaultNamespace",
      "contentTypes",
      "corePropertiesPath",
      "primaryPartPath",
      "conformance",
      "dialect",
      "application",
      "fonts",
    ],
    "document",
    "options",
  );
  if (options.rawParts?.length) {
    reject("document", "options", "rawParts", "RTF cannot retain opaque OOXML parts");
  }
  if (options.passthroughRelationships?.length) {
    reject(
      "document",
      "options",
      "passthroughRelationships",
      "RTF cannot retain OOXML relationship parts",
    );
  }
  const references: NumberingReferences = new Map();
  const context = createContext();
  context.footnotes = options.footnotes;
  context.endnotes = options.endnotes;
  for (const [index, font] of (options.fonts ?? []).entries()) {
    assertAllowed(font, ["name", "family", "characterSet"], "document", `fonts[${index}]`);
    if (typeof font.name !== "string" || font.name === "") {
      reject("document", `fonts[${index}]`, "name", "RTF requires a concrete font name");
    }
    registerFont(context, font);
  }
  const first = options.sections[0] ?? { children: [] };
  const body = [
    writeSettings(options.settings, "settings"),
    writeMetadata(options),
    assertTimestampType(options.createdType, "created"),
    assertTimestampType(options.modifiedType, "modified"),
    options.created === undefined || options.created === null
      ? ""
      : writeTimestamp("creatim", options.created, "created"),
    options.modified === undefined || options.modified === null
      ? ""
      : writeTimestamp("revtim", options.modified, "modified"),
    options.lastPrinted === undefined
      ? ""
      : writeTimestamp("printim", options.lastPrinted, "lastPrinted"),
    writeStyles(options.styles, context),
    writeNumberingTable(options.numbering, references),
    writeDestinations(first, "sections[0]"),
    writeSections(options, references, context),
  ].join("");
  const tables = `${writeFontTable(context)}${writeColorTable(context)}`;
  return `{\\rtf1${writeApplication(options)}${writeDefaultLanguage(options)}${tables}${body}}`;
}
