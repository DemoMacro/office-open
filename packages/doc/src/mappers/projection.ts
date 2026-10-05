import type {
  ParagraphChild,
  ParagraphOptions,
  SectionChild,
  TrackChangeChild,
} from "@office-open/docx";

import { DocParseError } from "../errors";
import type { CharacterRange, ParagraphRange, TextCharacter } from "./models";
import { paragraphPropertiesFor } from "./paragraph-projection";

const LEGACY_REVISION_AUTHOR = "Unknown";
const LEGACY_REVISION_DATE = "1970-01-01T00:00:00Z";

export function projectChildren(
  characters: TextCharacter[],
  ranges: readonly CharacterRange[],
  paragraphRanges: readonly ParagraphRange[] = [],
): SectionChild[] {
  const children: SectionChild[] = [];
  let paragraph: TextCharacter[] = [];

  const flushParagraph = (): void => {
    if (paragraph.length === 0) {
      children.push({
        paragraph:
          paragraphRanges.length > 0
            ? { children: [], ...paragraphPropertiesFor(paragraphRanges, paragraph[0]?.cp) }
            : "",
      });
      return;
    }
    const boundaries = new Set<number>([0, paragraph.length]);
    for (let index = 0; index < paragraph.length; index++) {
      const cp = paragraph[index]!.cp;
      for (const range of ranges) {
        if (cp === range.cpStart) boundaries.add(index);
      }
      if (index > 0 && paragraph[index]!.hyperlink !== paragraph[index - 1]!.hyperlink) {
        boundaries.add(index);
      }
      if (
        index > 0 &&
        paragraph[index]!.revision?.inserted !== paragraph[index - 1]!.revision?.inserted
      ) {
        boundaries.add(index);
      }
    }
    const sortedBoundaries = [...boundaries].sort((left, right) => left - right);
    const runs: ParagraphChild[] = [];
    for (let index = 0; index < sortedBoundaries.length - 1; index++) {
      const start = sortedBoundaries[index]!;
      const end = sortedBoundaries[index + 1]!;
      const text = paragraph
        .slice(start, end)
        .map((character) => character.value)
        .join("");
      if (text.length === 0) continue;
      const cp = paragraph[start]!.cp;
      const range = ranges.find((candidate) => candidate.cpStart <= cp && cp < candidate.cpEnd);
      const hyperlink = paragraph[start]!.hyperlink;
      const revision = paragraph[start]!.revision;
      const breakCount = text.split("\n").length - 1;
      const visibleText = text.replaceAll("\n", "");
      const revisionChildren: ParagraphChild[] = [];
      if (breakCount > 0) revisionChildren.push({ break: breakCount });
      if (visibleText.length > 0 && hyperlink) {
        revisionChildren.push({
          hyperlink: { ...hyperlink },
          children: [{ text: visibleText, ...range?.properties }],
        });
      } else if (visibleText.length > 0) {
        revisionChildren.push({ text: visibleText, ...range?.properties });
      }
      if (revision) {
        const children = revisionChildren.flatMap((child): TrackChangeChild[] =>
          "text" in child ? [child] : [],
        );
        const change = {
          author: LEGACY_REVISION_AUTHOR,
          date: LEGACY_REVISION_DATE,
          children,
        };
        runs.push(revision.inserted ? { insertion: change } : { deletion: change });
      } else {
        runs.push(...revisionChildren);
      }
    }
    const plainText = paragraph.map((character) => character.value).join("");
    const hasProperties = runs.some((run) =>
      Object.keys(run).some((key) => key !== "text" && key !== "break"),
    );
    const paragraphProperties = paragraphPropertiesFor(paragraphRanges, paragraph[0]?.cp);
    children.push({
      paragraph:
        hasProperties || runs.length > 1 || Object.keys(paragraphProperties).length > 0
          ? { children: runs, ...paragraphProperties }
          : plainText,
    });
    paragraph = [];
  };

  for (const character of characters) {
    if (character.paragraphEnd) flushParagraph();
    else paragraph.push(character);
  }
  if (paragraph.length > 0) flushParagraph();
  if (children.length === 0) children.push({ paragraph: "" });
  return children;
}

export function projectStory(
  characters: readonly TextCharacter[],
  ranges: readonly CharacterRange[],
  start: number,
  end: number,
  paragraphRanges: readonly ParagraphRange[] = [],
): SectionChild[] {
  return projectChildren(
    characters.filter((character) => character.cp >= start && character.cp < end),
    ranges,
    paragraphRanges,
  );
}

export function applyHyperlinkFields(
  characters: TextCharacter[],
  fields: readonly { start: number; end: number; instruction: string }[],
  rawCharacters: readonly string[],
): void {
  for (const field of fields) {
    const match = /^HYPERLINK\s+(?:"([^"]*)"|(\S+))/i.exec(field.instruction.trim());
    if (!match) continue;
    const target = match[2] ?? match[1];
    if (!target) continue;
    const hyperlink = target.startsWith("#") ? { anchor: target.slice(1) } : { url: target };
    for (
      let cp = field.end + 1;
      cp < rawCharacters.length && rawCharacters[cp] !== "\x15";
      cp += 1
    ) {
      const character = characters.find((candidate) => candidate.cp === cp);
      if (character) character.hyperlink = hyperlink;
    }
  }
}

export function parseOptional<T>(parse: () => T, fallback: T): T {
  try {
    return parse();
  } catch (error) {
    if (!(error instanceof DocParseError)) throw error;
    return fallback;
  }
}

export function projectCommentStory(
  characters: readonly TextCharacter[],
  ranges: readonly CharacterRange[],
  start: number,
  end: number,
): Array<string | ParagraphOptions> {
  return projectStory(characters, ranges, start, end).flatMap((child) =>
    "paragraph" in child ? [child.paragraph] : [],
  );
}
