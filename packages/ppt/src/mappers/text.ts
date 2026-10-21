import type {
  ParagraphDescriptorOptions,
  TextBodyOptions,
  TextBreakOptions,
  TextRunOptions,
} from "@office-open/core";
import type { ShapeOptions } from "@office-open/pptx";

import { PptParseError } from "../errors";
import {
  collectDescendants,
  decodeAnsi,
  decodeUtf16,
  findDirect,
  readInt32,
  recordBody,
  RecordType,
  type RecordNode,
} from "../records";
import { EMU_PER_INCH, MASTER_UNITS_PER_INCH } from "./constants";
import type { LegacyTextStyle, LegacyTextHyperlink, TextEntry } from "./models";
import { parseTextStyle } from "./text-style";

type PlaceholderType = NonNullable<ShapeOptions["placeholder"]>;
export function readHyperlinkTargets(
  view: DataView,
  document: RecordNode,
): ReadonlyMap<number, { url?: string; tooltip?: string }> {
  const targets = new Map<number, { url?: string; tooltip?: string }>();
  for (const list of collectDescendants(document, RecordType.externalObjectList)) {
    for (const link of list.children.filter((child) => child.type === RecordType.hyperlink)) {
      const atom = findDirect(link, RecordType.hyperlinkAtom);
      const strings = link.children.filter((child) => child.type === RecordType.interactiveString);
      const id = atom && atom.length >= 4 ? readInt32(view, atom, 0) : undefined;
      const url = strings[1] ? decodeUtf16(recordBody(view, strings[1])) : undefined;
      const tooltip = strings[0] ? decodeUtf16(recordBody(view, strings[0])) : undefined;
      if (id === undefined) {
        throw new PptParseError("Corrupt legacy PowerPoint hyperlink: missing target id", {
          part: "record",
          recordType: link.type,
          recordName: "HyperlinkContainer",
          reason: "missing-hyperlink-target",
        });
      }
      targets.set(id, { url, tooltip });
    }
  }
  return targets;
}

export function applyClientHyperlinks(
  body: TextBodyOptions,
  view: DataView,
  clientTextbox: RecordNode,
  targets: ReadonlyMap<number, { url?: string; tooltip?: string }>,
): TextBodyOptions {
  const links: LegacyTextHyperlink[] = [];
  for (const [index, record] of clientTextbox.children.entries()) {
    if (record.type !== RecordType.interactiveInfo) continue;
    const atom = findDirect(record, RecordType.interactiveInfoAtom);
    const range = clientTextbox.children[index + 1];
    if (
      !atom ||
      atom.length < 8 ||
      !range ||
      range.type !== RecordType.interactiveText ||
      range.length < 8
    )
      continue;
    const target = targets.get(readInt32(view, atom, 4));
    if (!target?.url) continue;
    links.push({
      start: readInt32(view, range, 0),
      end: readInt32(view, range, 4),
      target,
    });
  }
  if (links.length === 0) return body;

  let position = 0;
  for (const paragraph of body.paragraphs ?? []) {
    if (typeof paragraph === "string") continue;
    for (const run of paragraph.children ?? []) {
      if (typeof run === "string" || (!("text" in run) && !("hyperlink" in run))) continue;
      const textRun = run as TextRunOptions;
      const start = position;
      const end = position + (textRun.text?.length ?? 0);
      const link = links.find((candidate) => start < candidate.end && end > candidate.start);
      if (link) {
        textRun.hyperlink = { url: link.target.url, tooltip: link.target.tooltip };
      }
      position = end;
    }
    position += 1;
  }
  return body;
}

export function collectEmbeddedTextEntries(
  view: DataView,
  parent: RecordNode,
  fonts: ReadonlyMap<number, string>,
): TextEntry[] {
  const text: TextEntry[] = [];
  let placeholderType: number | undefined;
  let placeholderSize: "half" | "quarter" | undefined;
  let value: string | undefined;
  let hasDateTime = false;
  let hasSlideNumber = false;

  let pendingStyle: RecordNode | undefined;
  const flush = (): void => {
    if (value === undefined) return;
    if (value === "*" && hasSlideNumber) {
      text.push({ placeholderType: -3, text: "1" });
    } else if (value === "*" && hasDateTime) {
      text.push({ placeholderType: -2, text: "1/1/1" });
    } else {
      const entry = {
        placeholderType: placeholderType ?? -1,
        ...(placeholderSize !== undefined ? { placeholderSize } : {}),
        text: value,
        ...(pendingStyle
          ? { styles: parseTextStyle(view, pendingStyle, value.length, fonts) }
          : {}),
      };
      text.push(entry);
    }
    pendingStyle = undefined;
    value = undefined;
  };

  const flattened = flatten(parent);
  for (const child of flattened) {
    if (child.type === RecordType.textHeader) {
      flush();
      hasDateTime = false;
      hasSlideNumber = false;
      placeholderType = readInt32(view, child, 0);
      placeholderSize =
        placeholderType === 7 ? "half" : placeholderType === 8 ? "quarter" : undefined;
    } else if (placeholderType !== undefined && child.type === RecordType.textChars) {
      value = decodeUtf16(recordBody(view, child));
    } else if (placeholderType !== undefined && child.type === RecordType.textBytes) {
      value = decodeAnsi(recordBody(view, child));
    } else if (child.type === RecordType.styleTextPropAtom) {
      pendingStyle = child;
      flush();
    } else if (child.type === RecordType.dateTimeMetaCharAtom) {
      hasDateTime = true;
    } else if (child.type === RecordType.slideNumberAtom) {
      hasSlideNumber = true;
    }
  }
  flush();
  return text;
}

export function collectEmbeddedText(
  view: DataView,
  parent: RecordNode,
  fonts: ReadonlyMap<number, string>,
): string[] {
  return collectEmbeddedTextEntries(view, parent, fonts).map((entry) => entry.text);
}

export function legacyPlaceholder(
  placeholderType: number | undefined,
): PlaceholderType | undefined {
  if (placeholderType === 0) return "title";
  if (placeholderType === -2) return "date";
  if (placeholderType === -3) return "slideNumber";
  if (placeholderType === 1 || placeholderType === 5 || placeholderType === 7) return "body";
  if (placeholderType === 6) return "centerTitle";
  if (placeholderType === 2) return "body";
  return undefined;
}

export function flatten(record: RecordNode): RecordNode[] {
  const output: RecordNode[] = [];
  const visit = (node: RecordNode): void => {
    for (const child of node.children) {
      output.push(child);
      visit(child);
    }
  };
  visit(record);
  return output;
}

export function createTextBody(text: readonly (string | TextEntry)[]): TextBodyOptions {
  const paragraphs = text.flatMap((entry) =>
    typeof entry === "string"
      ? createPlainParagraphs(entry)
      : createStyledParagraphs(entry.text, entry.styles),
  );
  return paragraphs.length > 0 ? { paragraphs } : {};
}

function createPlainParagraphs(value: string): ParagraphDescriptorOptions[] {
  return splitParagraphs(value).map((paragraph) => ({ children: [{ text: paragraph }] }));
}

interface ParagraphSlice {
  readonly text: string;
  readonly start: number;
  readonly end: number;
}

function paragraphSlices(value: string): ParagraphSlice[] {
  const slices: ParagraphSlice[] = [];
  let start = 0;
  for (let index = 0; index < value.length; index += 1) {
    if (value[index] !== "\r") continue;
    slices.push({ text: value.slice(start, index), start, end: index });
    index += value[index + 1] === "\n" ? 1 : 0;
    start = index + 1;
  }
  if (start < value.length || slices.length === 0) {
    slices.push({ text: value.slice(start), start, end: value.length });
  }
  return slices.map((slice) => ({
    ...slice,
    text: slice.text.replace(/[\n\v]/g, "\n"),
  }));
}

function styleRuns(
  text: string,
  style: LegacyTextStyle["characters"][number] | undefined,
): (TextRunOptions | TextBreakOptions)[] {
  const output: (TextRunOptions | TextBreakOptions)[] = [];
  let segmentStart = 0;
  for (const [index, character] of text.split("").entries()) {
    if (character !== "\n" && character !== "\v") continue;
    if (index > segmentStart) {
      output.push({ ...style?.character, text: text.slice(segmentStart, index) });
    }
    output.push({ break: true, ...(style?.character ? { properties: style.character } : {}) });
    segmentStart = index + 1;
  }
  if (segmentStart < text.length) {
    output.push({ ...style?.character, text: text.slice(segmentStart) });
  }
  return output;
}

function createStyledParagraphs(
  value: string,
  styles: LegacyTextStyle | undefined,
): ParagraphDescriptorOptions[] {
  if (!styles) return createPlainParagraphs(value);

  return paragraphSlices(value).map((paragraph) => {
    const paragraphStyle = styles.paragraphs.find((style) => paragraph.start < style.characters);
    let styleStart = 0;
    const runs = styles.characters.flatMap((style) => {
      const styleEnd = styleStart + style.characters;
      const start = Math.max(styleStart, paragraph.start);
      const end = Math.min(styleEnd, paragraph.end);
      const result =
        start >= end
          ? []
          : styleRuns(paragraph.text.slice(start - paragraph.start, end - paragraph.start), style);
      styleStart = styleEnd;
      return result;
    });
    return {
      ...(paragraphStyle?.paragraph ? { properties: paragraphStyle.paragraph } : {}),
      children: runs.length > 0 ? runs : [{ text: paragraph.text }],
    };
  });
}

export function splitParagraphs(value: string): string[] {
  return value.split(/\r\n|\r/).map((paragraph) => paragraph.replace(/[\n\v]/g, "\n"));
}

export function masterUnitsToEmu(value: number): number {
  return Math.round((value * EMU_PER_INCH) / MASTER_UNITS_PER_INCH);
}
