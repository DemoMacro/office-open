import { CompoundFileReader } from "@office-open/core";
import type { DocumentOptions, SectionChild } from "@office-open/docx";

import { readSummaryInformation } from "../cfb/container";
import { decryptWordStreams } from "../cfb/encryption";
import { DocParseError } from "../errors";
import type { LegacyRevisionRange } from "../records/models";
import {
  parseBookmarks,
  parseBodyPictureAnchors,
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
import { parseCharacterFkp, parseParagraphFkp } from "./fkp";
import { legacyChildren, parseLegacyFib, parseLegacyText } from "./legacy-text";
import { noOpFibTableRecord } from "./no-op-records";
import { applyTables } from "./paragraph-projection";
import {
  applyHyperlinkFields,
  parseRequired,
  parseTolerant,
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
  const picturesByShapeId =
    fib.drawing.length === 0
      ? new Map()
      : parseRequired(
          { part: "data", path: "Escher/DggInfo", recordType: 50, recordName: "fcDggInfo" },
          () =>
            parsePictures(
              table.subarray(fib.drawing.offset, fib.drawing.offset + fib.drawing.length),
              dataStream ?? new Uint8Array(),
            ),
        );
  const shapeIdByCp =
    fib.bodyShapes.length === 0
      ? new Map()
      : parseRequired({ part: "table", path: "PlcspaMom", recordName: "fcPlcspaMom" }, () =>
          parseBodyPictureAnchors(table, fib.bodyShapes),
        );
  const characters = projectText(pieces, word, fib.totalCharacters, picturesByShapeId, shapeIdByCp);
  const characterEntries = parseRequired(
    { part: "table", path: "PlcfbteChpx", recordType: 12, recordName: "fcPlcfbteChpx" },
    () => parseBinTable(table, fib.characterBinTable, "character"),
  );
  const characterRanges = characterEntries.flatMap((entry) =>
    parseRequired({ part: "stream", path: "WordDocument/FKP", recordName: "CHPX FKP" }, () =>
      parseCharacterFkp(word, entry.page, pieces, fib.totalCharacters),
    ),
  );
  const paragraphRanges = parseRequired(
    { part: "table", path: "PlcfbtePapx", recordType: 13, recordName: "fcPlcfbtePapx" },
    () => parseBinTable(table, fib.paragraphBinTable, "paragraph"),
  ).flatMap((entry) =>
    parseRequired({ part: "stream", path: "WordDocument/FKP", recordName: "PAPX FKP" }, () =>
      parseParagraphFkp(word, entry.page, pieces, fib.totalCharacters),
    ),
  );

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
  const sections =
    fib.sectionTable.length === 0
      ? []
      : parseTolerant(
          { part: "table", path: "PlcfSed", recordType: 6, recordName: "fcPlcfSed" },
          [],
          () => parseSectionProperties(table, fib.sectionTable, word).sections,
        );
  const sectionModels: SectionModel[] = sections.length > 0 ? sections : [{}];
  const stylesheet: StylesheetModel =
    fib.styleSheet.length === 0
      ? { styles: [] }
      : parseTolerant(
          { part: "table", path: "Stshf", recordType: 1, recordName: "fcStshf" },
          { styles: [] },
          () => parseStylesheet(table, fib.styleSheet),
        );
  const bookmarks: ReturnType<typeof parseBookmarks> =
    fib.bookmarkNames.length === 0
      ? []
      : parseTolerant({ part: "table", path: "Bookmarks", recordName: "BKM" }, [], () =>
          parseBookmarks(table, fib.bookmarkNames, fib.bookmarkStarts, fib.bookmarkEnds),
        );
  const fields: ReturnType<typeof parseFields> =
    fib.fields.length === 0
      ? []
      : parseTolerant(
          { part: "table", path: "PlcffldMom", recordType: 16, recordName: "fcPlcffldMom" },
          [],
          () => parseFields(rawCharacters, table, fib.fields),
        );
  applyHyperlinkFields(characters, fields, rawCharacters);
  const fieldInstructionsByCp = new Map(
    fields
      .filter((field) => !/^HYPERLINK\b/i.test(field.instruction.trim()))
      .map((field) => [field.start, field.instruction]),
  );
  for (const character of characters) {
    if (character.complexField !== undefined) {
      character.complexField = fieldInstructionsByCp.get(character.cp) ?? "";
    }
  }
  const revisions: LegacyRevisionRange[] =
    fib.comments.length === 0
      ? []
      : parseTolerant(
          { part: "table", path: "PlcfandRef", recordType: 4, recordName: "fcPlcfandRef" },
          [],
          () =>
            parseCommentRanges(table, fib.comments).map((comment) => ({
              start: comment.start,
              end: comment.end,
              inserted: (comment.data[1]! & 0x01) === 0,
            })),
        );
  for (const revision of revisions) {
    for (const character of characters) {
      if (character.cp >= revision.start && character.cp < revision.end) {
        character.revision = { inserted: revision.inserted };
      }
    }
  }
  const lists: ReturnType<typeof parseLists> | undefined =
    fib.list.length === 0 && fib.listOverrides.length === 0
      ? undefined
      : parseTolerant({ part: "table", path: "Lists", recordName: "PLF" }, undefined, () =>
          parseLists(table, fib.list, fib.listOverrides),
        );
  const footnoteStart = ranges[1]!.start;
  const headerStart = ranges[2]!.start;
  const commentStart = headerStart + fib.ccpHeaders;
  const endnoteStart = commentStart + fib.ccpComments;
  const textboxStart = endnoteStart + fib.ccpEndnotes;
  const headerTextboxStart = textboxStart + fib.ccpTextboxes;

  const sectionChildGroups = sectionModels.map((sectionModel, index) => {
    const start = sectionModel.mainStart ?? sectionModels[index - 1]?.mainEnd ?? 0;
    const end = sectionModel.mainEnd ?? fib.ccpText;
    const projected = projectStory(
      characters,
      characterRanges,
      start,
      Math.min(end, fib.ccpText),
      paragraphRanges,
    );
    return applyBookmarks(
      applyTables(projected, paragraphRanges, characters),
      bookmarks,
      characters,
    );
  });
  const headerStories =
    fib.ccpHeaders === 0 || noOpFibTableRecord(11, fib.headerTable)
      ? []
      : parseTolerant(
          { part: "table", path: "PlcfHdd", recordType: 11, recordName: "fcPlcfHdd" },
          [],
          () => parseHeaderStreams(table, fib.headerTable, fib.ccpHeaders),
        );
  const headerStoryChildren = (sectionIndex: number, groupSlot: number): SectionChild[] => {
    const story = headerStories[6 + sectionIndex * 6 + groupSlot];
    if (!story) return [];
    return projectStory(
      characters,
      characterRanges,
      headerStart + story.start,
      headerStart + story.end,
    ).filter((child) => !("paragraph" in child && child.paragraph === ""));
  };
  const textboxBoundaries =
    fib.ccpTextboxes > 0 && !noOpFibTableRecord(56, fib.textboxTable)
      ? parseTolerant(
          { part: "table", path: "PlcftxbxTxt", recordType: 56, recordName: "fcPlcftxbxTxt" },
          [],
          () => parseTextboxBoundaries(table, fib.textboxTable, fib.ccpTextboxes),
        )
      : [];
  const anchorCpByShapeId = new Map(
    [...shapeIdByCp].map(([characterPosition, shapeId]) => [shapeId, characterPosition]),
  );
  for (const boundary of textboxBoundaries) {
    const anchorCp = anchorCpByShapeId.get(boundary.shapeId);
    if (anchorCp === undefined) {
      throw new DocParseError("Invalid DOC textbox anchor: shape is not anchored in main text", {
        part: "table",
        path: "PlcftxbxTxt/FTXBXS/lid",
        shapeId: boundary.shapeId,
        reason: "missing-anchor",
      });
    }
    const sectionIndex = sectionModels.findIndex((sectionModel, index) => {
      const start = sectionModel.mainStart ?? sectionModels[index - 1]?.mainEnd ?? 0;
      const end = sectionModel.mainEnd ?? fib.ccpText;
      return anchorCp >= start && anchorCp < end;
    });
    if (sectionIndex === -1) {
      throw new DocParseError("Invalid DOC textbox anchor: main-text position has no section", {
        part: "table",
        path: "PlcspaMom",
        shapeId: boundary.shapeId,
        characterPosition: anchorCp,
        reason: "unmapped-section",
      });
    }
    sectionChildGroups[sectionIndex]?.push({
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
    fib.ccpHeaderTextboxes > 0 && !noOpFibTableRecord(58, fib.headerTextboxTable)
      ? parseTolerant(
          {
            part: "table",
            path: "PlcfHdrTxbxTxt",
            recordType: 58,
            recordName: "fcPlcfHdrTxbxTxt",
          },
          [],
          () => parseTextboxBoundaries(table, fib.headerTextboxTable, fib.ccpHeaderTextboxes),
        )
      : [];
  const sectionHeaderStories = sectionModels.map((_, sectionIndex) => ({
    defaultHeader: headerStoryChildren(sectionIndex, 1),
    evenHeader: headerStoryChildren(sectionIndex, 0),
    firstHeader: headerStoryChildren(sectionIndex, 4),
    defaultFooter: headerStoryChildren(sectionIndex, 3),
    evenFooter: headerStoryChildren(sectionIndex, 2),
    firstFooter: headerStoryChildren(sectionIndex, 5),
  }));
  for (const boundary of headerTextboxBoundaries) {
    sectionHeaderStories[0]?.defaultHeader.push({
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

  const result: DocumentOptions = {
    sections: sectionModels.map((sectionModel, index) => {
      const hasPageSize = sectionModel.width !== undefined && sectionModel.height !== undefined;
      const hasPageMargin =
        sectionModel.marginTop !== undefined ||
        sectionModel.marginRight !== undefined ||
        sectionModel.marginBottom !== undefined ||
        sectionModel.marginLeft !== undefined;
      return {
        children: sectionChildGroups[index] ?? [],
        ...(hasPageSize || hasPageMargin
          ? {
              properties: {
                ...(hasPageSize
                  ? {
                      pageSize: {
                        width: sectionModel.width!,
                        height: sectionModel.height!,
                        ...(sectionModel.orientation
                          ? { orientation: sectionModel.orientation }
                          : {}),
                      },
                    }
                  : {}),
                ...(hasPageMargin
                  ? {
                      pageMargin: {
                        ...(sectionModel.marginTop !== undefined
                          ? { top: sectionModel.marginTop }
                          : {}),
                        ...(sectionModel.marginRight !== undefined
                          ? { right: sectionModel.marginRight }
                          : {}),
                        ...(sectionModel.marginBottom !== undefined
                          ? { bottom: sectionModel.marginBottom }
                          : {}),
                        ...(sectionModel.marginLeft !== undefined
                          ? { left: sectionModel.marginLeft }
                          : {}),
                      },
                    }
                  : {}),
              },
            }
          : {}),
        ...(sectionHeaderStories[index] &&
        (sectionHeaderStories[index]!.defaultHeader.length > 0 ||
          sectionHeaderStories[index]!.evenHeader.length > 0 ||
          sectionHeaderStories[index]!.firstHeader.length > 0)
          ? {
              headers: {
                ...(sectionHeaderStories[index]!.defaultHeader.length > 0
                  ? { default: sectionHeaderStories[index]!.defaultHeader }
                  : {}),
                ...(sectionHeaderStories[index]!.evenHeader.length > 0
                  ? { even: sectionHeaderStories[index]!.evenHeader }
                  : {}),
                ...(sectionHeaderStories[index]!.firstHeader.length > 0
                  ? { first: sectionHeaderStories[index]!.firstHeader }
                  : {}),
              },
            }
          : {}),
        ...(sectionHeaderStories[index] &&
        (sectionHeaderStories[index]!.defaultFooter.length > 0 ||
          sectionHeaderStories[index]!.evenFooter.length > 0 ||
          sectionHeaderStories[index]!.firstFooter.length > 0)
          ? {
              footers: {
                ...(sectionHeaderStories[index]!.defaultFooter.length > 0
                  ? { default: sectionHeaderStories[index]!.defaultFooter }
                  : {}),
                ...(sectionHeaderStories[index]!.evenFooter.length > 0
                  ? { even: sectionHeaderStories[index]!.evenFooter }
                  : {}),
                ...(sectionHeaderStories[index]!.firstFooter.length > 0
                  ? { first: sectionHeaderStories[index]!.firstFooter }
                  : {}),
              },
            }
          : {}),
      };
    }),
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
          footnotes: parseRequired(
            { part: "table", path: "PlcffndTxt", recordType: 3, recordName: "fcPlcffndTxt" },
            () => parseNoteBoundaries(table, fib.footnoteTable, fib.ccpFootnotes),
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
          endnotes: parseRequired(
            { part: "table", path: "PlcfendTxt", recordType: 47, recordName: "fcPlcfendTxt" },
            () => parseNoteBoundaries(table, fib.endnoteTable, fib.ccpEndnotes),
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
  return result;
}
