import type { DocumentOptions } from "@office-open/docx";
import {
  attributeNumber,
  attributeString,
  childNamed,
  childrenNamed,
  readXml,
  xmlElement,
  type OdfFiles,
} from "@office-open/odf";

import {
  NUM_FORMAT_DOCX,
  NUM_FORMAT_ODF,
  type EndnoteProperties,
  type FootnoteProperties,
} from "./section";

/** Footnote/endnote numbering config → ODF text:notes-configuration elements. */
export function notesConfigurationXml(settings: DocumentOptions["settings"]): string {
  return (
    notesConfigXml("footnote", settings?.footnoteProperties) +
    notesConfigXml("endnote", settings?.endnoteProperties)
  );
}

export function notesConfigXml(
  noteClass: "footnote" | "endnote",
  properties: FootnoteProperties | EndnoteProperties | undefined,
): string {
  if (!properties) return "";
  const numFmt =
    properties.format ?? (properties.numFmt ? NUM_FORMAT_ODF[properties.numFmt] : undefined);
  return xmlElement("text:notes-configuration", {
    "text:note-class": noteClass,
    "style:num-format": numFmt,
    "text:start-value": properties.numStart,
    "text:start-numbering-at":
      properties.numRestart === "continuous"
        ? "document"
        : properties.numRestart === "eachSect"
          ? "chapter"
          : properties.numRestart === "eachPage"
            ? "page"
            : undefined,
    "text:footnotes-position":
      properties.pos === "pageBottom"
        ? "page"
        : properties.pos === "beneathText"
          ? "text"
          : properties.pos === "sectEnd"
            ? "section"
            : properties.pos === "docEnd"
              ? "document"
              : undefined,
  });
}

/** ODF notes-configuration → docx footnote/endnote document properties. */
export function parseNotesConfiguration(files: OdfFiles): {
  footnoteProperties?: FootnoteProperties;
  endnoteProperties?: EndnoteProperties;
} {
  const styles = readXml(files, "styles.xml");
  const result: { footnoteProperties?: FootnoteProperties; endnoteProperties?: EndnoteProperties } =
    {};
  for (const config of childrenNamed(
    childNamed(styles, "office:styles"),
    "text:notes-configuration",
  )) {
    const noteClass = attributeString(config, "text:note-class");
    if (noteClass !== "footnote" && noteClass !== "endnote") continue;
    const rawNumFormat = attributeString(config, "style:num-format");
    const mappedNumFormat = rawNumFormat ? NUM_FORMAT_DOCX[rawNumFormat] : undefined;
    const base = {
      pos: docxNotePosition(attributeString(config, "text:footnotes-position")),
      numFmt: mappedNumFormat,
      format: mappedNumFormat ? undefined : rawNumFormat,
      numStart: attributeNumber(config, "text:start-value"),
      numRestart: docxNumberRestart(attributeString(config, "text:start-numbering-at")),
    };
    if (noteClass === "footnote") result.footnoteProperties = base as FootnoteProperties;
    else result.endnoteProperties = base as EndnoteProperties;
  }
  return result;
}

export function docxNotePosition(value: string | undefined): FootnoteProperties["pos"] {
  if (value === "page") return "pageBottom";
  if (value === "text") return "beneathText";
  if (value === "section") return "sectEnd";
  if (value === "document") return "docEnd";
  return undefined;
}

export function docxNumberRestart(value: string | undefined): FootnoteProperties["numRestart"] {
  if (value === "document") return "continuous";
  if (value === "chapter") return "eachSect";
  if (value === "page") return "eachPage";
  return undefined;
}
