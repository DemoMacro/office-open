import { CompoundFileReader } from "@office-open/core";
import type { DocumentOptions, ParagraphChild, SectionChild } from "@office-open/docx";

import { readSummaryInformation } from "../cfb/container";
import { decryptWordStreams } from "../cfb/encryption";
import { DocParseError } from "../errors";
import type { LegacyRevisionRange } from "../records/models";
import {
  parseBookmarks,
  parseFields,
  parseLists,
  parsePictures,
  parseSectionProperties,
  parseStylesheet,
  type SectionModel,
  type StylesheetModel,
} from "../records/structures";
import { parsePieceTable } from "../streams/pieces";
import { readUint16 } from "./binary";
import { applyBookmarks } from "./bookmarks";
import { parseCommentRanges } from "./comment-ranges";
import { FLAG_ENCRYPTED, FLAG_TABLE_ONE } from "./constants";
import { parseFib } from "./fib";
import { overlappingPieces, parseCharacterFkp, parseParagraphFkp } from "./fkp";
import { legacyChildren, parseLegacyFib, parseLegacyText } from "./legacy-text";
import { applyTables } from "./paragraph-projection";
import {
  applyHyperlinkFields,
  parseOptional,
  projectCommentStory,
  projectStory,
} from "./projection";
import {
  parseBinTable,
  parseHeaderStreams,
  parseNoteBoundaries,
  parseTextboxBoundaries,
} from "./stream-tables";
import { decodePiece, projectText } from "./text";

export { parseCommentRanges };

export function parseInternal(data: Uint8Array, password?: string): DocumentOptions {
  let reader: CompoundFileReader;
  try {
    reader = new CompoundFileReader(data);
  } catch (error) {
    throw new DocParseError(
      `Input is not a supported Compound File Binary document: ${(error as Error).message}`,
      { part: "container", path: "/", reason: "invalid-container" },
    );
  }
  if (!reader.entry("WordDocument")) {
    throw new DocParseError("Invalid Word document: WordDocument stream is missing", {
      part: "stream",
      path: "WordDocument",
      reason: "missing-required-stream",
    });
  }
  let word = reader.read("WordDocument");
  const legacy = readUint16(word, 0, "Invalid Word document: FIB is truncated") === 0xa5dc;
  if (legacy) {
    const fib = parseLegacyFib(word);
    return { sections: [{ children: legacyChildren(parseLegacyText(word, fib)) }] };
  }
  const flags = readUint16(word, 10, "Invalid Word document: FIB is truncated");
  let tablePath = (flags & FLAG_TABLE_ONE) !== 0 ? "1Table" : "0Table";
  let table: Uint8Array | undefined = reader.entry(tablePath) ? reader.read(tablePath) : undefined;
  let dataStream = reader.entry("Data") ? reader.read("Data") : undefined;
  if ((flags & FLAG_ENCRYPTED) !== 0) {
    if (!table) {
      throw new DocParseError("Encrypted Word documents are not supported", {
        part: "table",
        path: tablePath,
        reason: "encrypted-unsupported",
      });
    }
    const decrypted = decryptWordStreams(word, table, reader, password);
    word = decrypted.word;
    table = decrypted.table;
    tablePath = decrypted.tablePath;
    dataStream = decrypted.data;
  }
  const fib = parseFib(word);
  const metadata = readSummaryInformation(reader);
  if (!table) {
    throw new DocParseError(`Invalid Word document: ${tablePath} stream is missing`, {
      part: "table",
      path: tablePath,
      reason: "missing-required-stream",
    });
  }
  if (fib.totalCharacters === 0) return { sections: [{ children: [{ paragraph: "" }] }] };

  const pieces = parsePieceTable(table, fib.clx);
  if (fib.totalCharacters > pieces[pieces.length - 1]!.cpEnd) {
    throw new DocParseError("Invalid Word document: main text extends beyond the piece table");
  }
  const characters = projectText(pieces, word, fib.totalCharacters);
  const characterEntries = parseOptional(
    () => parseBinTable(table, fib.characterBinTable, "character"),
    [],
  );
  const characterRanges = characterEntries.flatMap((entry) =>
    parseOptional(() => parseCharacterFkp(word, entry.page, pieces, fib.totalCharacters), []),
  );
  const paragraphRanges = parseOptional(
    () => parseBinTable(table, fib.paragraphBinTable, "paragraph"),
    [],
  )
    .flatMap((entry) => parseOptional(() => parseParagraphFkp(word, entry.page), []))
    .map((range) => {
      const overlaps = overlappingPieces(pieces, fib.totalCharacters, range.cpStart, range.cpEnd);
      if (overlaps.length === 0) return undefined;
      return {
        cpStart: Math.min(...overlaps.map((overlap) => overlap.cpStart)),
        cpEnd: Math.max(...overlaps.map((overlap) => overlap.cpEnd)),
        properties: range.properties,
      };
    })
    .filter((range) => range !== undefined);

  const ranges = [
    { start: 0, end: fib.ccpText },
    { start: fib.ccpText, end: fib.ccpText + fib.ccpFootnotes },
    {
      start: fib.ccpText + fib.ccpFootnotes,
      end: fib.ccpText + fib.ccpFootnotes + fib.ccpHeaders,
    },
  ];
  const rawCharacters: string[] = [];
  for (const piece of pieces) {
    const cpEnd = Math.min(piece.cpEnd, fib.totalCharacters);
    const characters = decodePiece(word, { ...piece, cpEnd });
    for (let index = 0; index < characters.length; index += 1) {
      rawCharacters.push(characters[index]!);
    }
  }
  let sections: SectionModel[] = [];
  if (fib.sectionTable.length > 0) {
    try {
      sections = parseSectionProperties(table, fib.sectionTable, word).sections;
    } catch {
      sections = [];
    }
  }
  const sectionModels: SectionModel[] = sections.length > 0 ? sections : [{}];
  let stylesheet: StylesheetModel = { styles: [] };
  if (fib.styleSheet.length > 0) {
    try {
      stylesheet = parseStylesheet(table, fib.styleSheet);
    } catch {
      stylesheet = { styles: [] };
    }
  }
  let bookmarks: ReturnType<typeof parseBookmarks> = [];
  if (fib.bookmarkNames.length > 0) {
    try {
      bookmarks = parseBookmarks(table, fib.bookmarkNames, fib.bookmarkStarts, fib.bookmarkEnds);
    } catch {
      bookmarks = [];
    }
  }
  let fields: ReturnType<typeof parseFields> = [];
  if (fib.fields.length > 0) {
    fields = parseOptional(() => parseFields(rawCharacters, table, fib.fields), []);
  }
  applyHyperlinkFields(characters, fields, rawCharacters);
  let revisions: LegacyRevisionRange[] = [];
  if (fib.comments.length > 0) {
    revisions = parseOptional(
      () =>
        parseCommentRanges(table, fib.comments).map((comment) => ({
          start: comment.start,
          end: comment.end,
          inserted: (comment.data[1]! & 0x01) === 0,
        })),
      [],
    );
  }
  for (const revision of revisions) {
    for (const character of characters) {
      if (character.cp >= revision.start && character.cp < revision.end) {
        character.revision = { inserted: revision.inserted };
      }
    }
  }
  const pictures =
    fib.drawing.length > 0 && dataStream
      ? parseOptional(
          () =>
            parsePictures(
              table.subarray(fib.drawing.offset, fib.drawing.offset + fib.drawing.length),
              dataStream,
            ),
          [],
        )
      : [];
  let lists: ReturnType<typeof parseLists> | undefined;
  if (fib.list.length > 0 || fib.listOverrides.length > 0) {
    try {
      lists = parseLists(table, fib.list, fib.listOverrides);
    } catch {
      lists = undefined;
    }
  }
  const footnoteStart = ranges[1]!.start;
  const headerStart = ranges[2]!.start;
  const commentStart = headerStart + fib.ccpHeaders;
  const endnoteStart = commentStart + fib.ccpComments;
  const textboxStart = endnoteStart + fib.ccpEndnotes;
  const headerTextboxStart = textboxStart + fib.ccpTextboxes;

  const projectedChildren = projectStory(
    characters,
    characterRanges,
    0,
    fib.ccpText,
    paragraphRanges,
  );
  const sectionChildren = applyTables(projectedChildren, paragraphRanges, characters);
  const sectionChildrenWithBookmarks = applyBookmarks(sectionChildren, bookmarks, characters);
  const headerStories = parseOptional(
    () => parseHeaderStreams(table, fib.headerTable, fib.ccpHeaders),
    [],
  );
  const headerStoryChildren = (index: number): SectionChild[] => {
    const story = headerStories[index];
    if (!story) return [];
    return projectStory(
      characters,
      characterRanges,
      headerStart + story.start,
      headerStart + story.end,
    ).filter((child) => !("paragraph" in child && child.paragraph === ""));
  };
  const defaultHeaderChildren = headerStoryChildren(7);
  const evenHeaderChildren = headerStoryChildren(6);
  const firstHeaderChildren = headerStoryChildren(10);
  const defaultFooterChildren = headerStoryChildren(9);
  const evenFooterChildren = headerStoryChildren(8);
  const firstFooterChildren = headerStoryChildren(11);
  const textboxBoundaries =
    fib.ccpTextboxes > 0
      ? parseOptional(() => parseTextboxBoundaries(table, fib.textboxTable, fib.ccpTextboxes), [])
      : [];
  for (const boundary of textboxBoundaries) {
    sectionChildren.push({
      textbox: {
        children: projectStory(
          characters,
          characterRanges,
          textboxStart + boundary.start,
          textboxStart + boundary.end,
        ),
      },
    });
  }
  const headerTextboxBoundaries =
    fib.ccpHeaderTextboxes > 0
      ? parseOptional(
          () => parseTextboxBoundaries(table, fib.headerTextboxTable, fib.ccpHeaderTextboxes),
          [],
        )
      : [];
  for (const boundary of headerTextboxBoundaries) {
    defaultHeaderChildren.push({
      textbox: {
        children: projectStory(
          characters,
          characterRanges,
          headerTextboxStart + boundary.start,
          headerTextboxStart + boundary.end,
        ),
      },
    });
  }

  const sectionGeometry = sectionModels[0] ?? {};
  const hasPageSize = sectionGeometry.width !== undefined && sectionGeometry.height !== undefined;
  const hasPageMargin =
    sectionGeometry.marginTop !== undefined ||
    sectionGeometry.marginRight !== undefined ||
    sectionGeometry.marginBottom !== undefined ||
    sectionGeometry.marginLeft !== undefined;
  const sectionProperties = {
    ...(hasPageSize
      ? {
          pageSize: {
            width: sectionGeometry.width!,
            height: sectionGeometry.height!,
            ...(sectionGeometry.orientation ? { orientation: sectionGeometry.orientation } : {}),
          },
        }
      : {}),
    ...(hasPageMargin
      ? {
          pageMargin: {
            ...(sectionGeometry.marginTop !== undefined ? { top: sectionGeometry.marginTop } : {}),
            ...(sectionGeometry.marginRight !== undefined
              ? { right: sectionGeometry.marginRight }
              : {}),
            ...(sectionGeometry.marginBottom !== undefined
              ? { bottom: sectionGeometry.marginBottom }
              : {}),
            ...(sectionGeometry.marginLeft !== undefined
              ? { left: sectionGeometry.marginLeft }
              : {}),
          },
        }
      : {}),
  };
  const result: DocumentOptions = {
    sections: [
      {
        children: sectionChildrenWithBookmarks,
        ...(Object.keys(sectionProperties).length > 0 ? { properties: sectionProperties } : {}),
        ...(defaultHeaderChildren.length > 0 ||
        evenHeaderChildren.length > 0 ||
        firstHeaderChildren.length > 0
          ? {
              headers: {
                ...(defaultHeaderChildren.length > 0 ? { default: defaultHeaderChildren } : {}),
                ...(evenHeaderChildren.length > 0 ? { even: evenHeaderChildren } : {}),
                ...(firstHeaderChildren.length > 0 ? { first: firstHeaderChildren } : {}),
              },
            }
          : {}),
        ...(defaultFooterChildren.length > 0 ||
        evenFooterChildren.length > 0 ||
        firstFooterChildren.length > 0
          ? {
              footers: {
                ...(defaultFooterChildren.length > 0 ? { default: defaultFooterChildren } : {}),
                ...(evenFooterChildren.length > 0 ? { even: evenFooterChildren } : {}),
                ...(firstFooterChildren.length > 0 ? { first: firstFooterChildren } : {}),
              },
            }
          : {}),
      },
    ],
    ...(stylesheet.styles.length > 0
      ? {
          styles: {
            paragraphStyles: stylesheet.styles
              .filter((style) => style.type === "paragraph")
              .map((style) => ({ id: style.id, name: style.name, basedOn: style.basedOn })),
            characterStyles: stylesheet.styles
              .filter((style) => style.type === "character")
              .map((style) => ({ id: style.id, name: style.name, basedOn: style.basedOn })),
          },
        }
      : {}),
    ...(lists && (lists.abstractNumberings.length > 0 || lists.overrides.length > 0)
      ? {
          numbering: {
            abstractNumberings: lists.abstractNumberings.map((numbering) => ({
              reference: numbering.reference,
              levels: numbering.levels,
            })),
          },
        }
      : {}),
    ...(fib.ccpFootnotes > 0
      ? {
          footnotes: parseOptional(
            () => parseNoteBoundaries(table, fib.footnoteTable, fib.ccpFootnotes),
            [],
          ).map((boundary, index) => ({
            id: index + 1,
            children: projectStory(
              characters,
              characterRanges,
              footnoteStart + boundary.start,
              footnoteStart + boundary.end,
            ),
          })),
        }
      : {}),
    ...(fib.ccpComments > 0
      ? {
          comments: [
            {
              id: 1,
              children: projectCommentStory(
                characters,
                characterRanges,
                commentStart,
                endnoteStart,
              ),
            },
          ],
        }
      : {}),
    ...(fib.ccpEndnotes > 0
      ? {
          endnotes: parseOptional(
            () => parseNoteBoundaries(table, fib.endnoteTable, fib.ccpEndnotes),
            [],
          ).map((boundary, index) => ({
            id: index + 1,
            children: projectStory(
              characters,
              characterRanges,
              endnoteStart + boundary.start,
              endnoteStart + boundary.end,
            ),
          })),
        }
      : {}),
  };
  Object.assign(result, metadata);
  const section = result.sections[0];
  if (section) {
    const canonicalFields = fields
      .filter((field) => !/^HYPERLINK\b/i.test(field.instruction.trim()))
      .map((field) => ({ complexField: { instruction: field.instruction } }));
    const canonicalPictures = pictures.map((picture): ParagraphChild => {
      if (picture.type === "unknown") {
        throw new DocParseError("Unknown DOC picture format has no canonical picture mapping");
      }
      const type = picture.type === "jpeg" ? "jpg" : picture.type === "tiff" ? "tif" : picture.type;
      return {
        picture: {
          type,
          data: picture.data,
          transformation: { width: picture.width, height: picture.height },
        },
      };
    });
    section.children = [
      ...canonicalPictures.map((picture) => ({ paragraph: { children: [picture] } })),
      ...canonicalFields.map((field) => ({ paragraph: { children: [field] } })),
      ...section.children,
    ];
  }
  return result;
}
