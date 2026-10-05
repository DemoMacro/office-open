import type { TextBodyOptions, TextRunOptions } from "@office-open/core";

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
import type { LegacyTextHyperlink } from "./models";
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
      if (id === undefined || !url || url.includes(",")) continue;
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

export function collectEmbeddedText(view: DataView, parent: RecordNode): string[] {
  const text: string[] = [];
  let placeholderType: number | undefined;
  let value: string | undefined;
  let hasDateTime = false;
  let hasSlideNumber = false;

  const flush = (): void => {
    if (value === undefined) return;
    if (value === "*" && hasSlideNumber) text.push("1");
    else if (value === "*" && hasDateTime) text.push("1/1/1");
    else text.push(value);
    value = undefined;
  };

  for (const child of flatten(parent)) {
    if (child.type === RecordType.textHeader) {
      flush();
      hasDateTime = false;
      hasSlideNumber = false;
      placeholderType = readInt32(view, child, 0);
    } else if (placeholderType !== undefined && child.type === RecordType.textChars) {
      value = decodeUtf16(recordBody(view, child));
    } else if (placeholderType !== undefined && child.type === RecordType.textBytes) {
      value = decodeAnsi(recordBody(view, child));
    } else if (child.type === RecordType.dateTimeAtom) {
      hasDateTime = true;
    } else if (child.type === RecordType.slideNumberAtom) {
      hasSlideNumber = true;
    }
  }
  flush();
  return text;
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

export function createTextBody(text: readonly string[]): TextBodyOptions {
  const paragraphs = text
    .flatMap((value) => splitParagraphs(value))
    .map((value) => ({ children: [{ text: value }] satisfies TextRunOptions[] }));
  return paragraphs.length > 0 ? { paragraphs } : {};
}

export function splitParagraphs(value: string): string[] {
  return value.split(/\r\n|\r/).map((paragraph) => paragraph.replace(/[\n\v]/g, "\n"));
}

export function masterUnitsToEmu(value: number): number {
  return Math.round((value * EMU_PER_INCH) / MASTER_UNITS_PER_INCH);
}
