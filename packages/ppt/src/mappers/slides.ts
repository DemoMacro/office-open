import type { AnimationEntry, MasterDefinition, SlideChild, SlideOptions } from "@office-open/pptx";

import { PptParseError } from "../errors";
import { readPersistedRecord } from "../parts/persist";
import {
  RecordType,
  collectDescendants,
  findDirect,
  readInt16,
  readInt32,
  readRecordHeader,
  readRecordTree,
  readUint32,
  type RecordNode,
} from "../records";
import { RECORD_BODY_OFFSET } from "./constants";
import type { DrawingContext, LegacyAnimation, SlideTextGroup } from "./models";
import { readDrawingChildren, textBodyToText } from "./shapes";
import { collectEmbeddedText, createTextBody, splitParagraphs } from "./text";
export function readSlide(
  view: DataView,
  references: Map<number, number>,
  group: SlideTextGroup,
  context: DrawingContext,
  defaultMaster: string | undefined,
  masterNames: ReadonlyMap<number, string>,
  notesReferences: ReadonlyMap<number, number>,
): SlideOptions {
  const offset = references.get(group.persistReference);
  if (offset === undefined) {
    throw new PptParseError("Corrupt legacy PowerPoint file: missing slide persist object");
  }
  const slideHeader = readRecordHeader(view, offset);
  const slide = readRecordTree(view, offset, slideHeader.end)[0];
  if (!slide || slide.type !== RecordType.slide) {
    throw new PptParseError("Corrupt legacy PowerPoint file: unexpected slide persist object");
  }
  const drawing = findDirect(slide, RecordType.ppDrawing);
  const drawingChildren = drawing ? readDrawingChildren(view, drawing, context) : [];
  const children = appendUnreferencedText(drawingChildren, context);
  const slideAtom = findDirect(slide, RecordType.slideAtom);
  const hasSlideAtom = slideAtom !== undefined && slideAtom.length >= 24;
  const masterId = hasSlideAtom ? readUint32(view, slideAtom!, 16) : undefined;
  const notesId = hasSlideAtom ? readUint32(view, slideAtom!, 20) : undefined;
  const master = masterId !== undefined && masterId > 0 ? masterNames.get(masterId) : undefined;
  const result: SlideOptions = {};
  if (children.length > 0) result.children = children;
  const slideMaster = master ?? defaultMaster;
  if (slideMaster) result.master = slideMaster;
  const notes = readNotes(view, notesReferences, notesId, context.fonts);
  if (notes) result.notes = notes;
  const animations = drawing
    ? readSlideAnimations(view, drawing).map((animation) => projectAnimation(view, animation))
    : [];
  if (animations.length > 0) result.animations = animations;
  return result;
}

export function readNotes(
  view: DataView,
  notesReferences: ReadonlyMap<number, number>,
  notesId: number | undefined,
  fonts: ReadonlyMap<number, string>,
): string | undefined {
  if (!notesId) return undefined;
  const offset = notesReferences.get(notesId);
  if (offset === undefined) {
    return undefined;
  }
  const notes = readPersistedRecord(view, offset, [RecordType.notes]);
  const drawing = findDirect(notes, RecordType.ppDrawing);
  if (!drawing) {
    throw new PptParseError("Corrupt legacy PowerPoint file: notes container has no drawing", {
      part: "record",
      recordType: RecordType.notes,
      recordName: "NotesContainer",
      notesId,
      offset,
      reason: "missing-notes-drawing",
    });
  }
  const paragraphs = collectEmbeddedText(view, drawing, fonts)
    .flatMap((value) => splitParagraphs(value))
    .filter((value) => value.length > 0);
  return paragraphs.length > 0 ? paragraphs.join("\n") : undefined;
}

export function readNotesPersistReferences(
  view: DataView,
  references: ReadonlyMap<number, number>,
): ReadonlyMap<number, number> {
  const notesReferences = new Map<number, number>();
  for (const offset of new Set(references.values())) {
    const header = readRecordHeader(view, offset);
    if (header.type !== RecordType.notes) continue;
    const notes = readPersistedRecord(view, offset, [RecordType.notes]);
    const atom = findDirect(notes, RecordType.notesAtom);
    if (!atom || atom.length < 4) {
      throw new PptParseError("Corrupt legacy PowerPoint notes container: missing notes id", {
        part: "record",
        recordType: RecordType.notesAtom,
        recordName: "NotesAtom",
        offset,
        reason: "missing-notes-id",
      });
    }
    const notesId = readUint32(view, atom, 0);
    const existingOffset = notesReferences.get(notesId);
    if (existingOffset !== undefined && existingOffset !== offset) {
      throw new PptParseError("Corrupt legacy PowerPoint file: duplicate notes id", {
        part: "record",
        recordType: RecordType.notesAtom,
        recordName: "NotesAtom",
        notesId,
        offset,
        reason: "duplicate-notes-id",
      });
    }
    notesReferences.set(notesId, offset);
  }
  return notesReferences;
}

export function readSlideAnimations(view: DataView, drawing: RecordNode): LegacyAnimation[] {
  const shapes = collectDescendants(drawing, RecordType.escherShapeContainer);
  const animations: LegacyAnimation[] = [];
  for (const shape of shapes) {
    try {
      const shapeRecord = findDirect(shape, RecordType.escherShape);
      const container = findDirect(shape, RecordType.animationInfo);
      const atom = container && findDirect(container, RecordType.animationInfoAtom);
      if (atom && atom.length >= 28) {
        animations.push({
          shapeId:
            shapeRecord && shapeRecord.length >= 4 ? readInt32(view, shapeRecord, 0) : undefined,
          atom,
        });
      }
    } catch {
      continue;
    }
  }
  return animations.sort(
    (first, second) => animationOrder(view, first) - animationOrder(view, second),
  );
}

function animationOrder(view: DataView, animation: LegacyAnimation): number {
  try {
    return readInt16(view, animation.atom, 16);
  } catch {
    return 0;
  }
}

function projectAnimation(view: DataView, animation: LegacyAnimation): AnimationEntry {
  const atom = animation.atom;
  const effect = view.getUint8(atom.offset + RECORD_BODY_OFFSET + 21);
  const mask = view.getUint16(atom.offset + RECORD_BODY_OFFSET + 4, true);
  const automatic = ((mask >>> 2) & 1) === 1;
  let delay = 0;
  try {
    delay = Math.max(0, readInt32(view, atom, 12));
  } catch {
    delay = 0;
  }
  return {
    type: animationType(effect),
    class: "entrance",
    trigger: automatic ? "afterPrevious" : "onClick",
    ...(delay > 0 ? { delay } : {}),
    ...(animation.shapeId !== undefined && animation.shapeId > 0
      ? { shapeId: animation.shapeId }
      : {}),
  };
}

function animationType(effect: number): AnimationEntry["type"] {
  if (effect === 0x02) return "blinds";
  if (effect === 0x03) return "checker";
  if (effect === 0x04) return "cover";
  if (effect === 0x05) return "dissolve";
  if (effect === 0x06) return "fade";
  if (effect === 0x08) return "randomBars";
  if (effect === 0x09) return "strips";
  if (effect === 0x0a) return "wipe";
  if (effect === 0x0b) return "fly";
  if (effect === 0x0d) return "split";
  if (effect === 0x1a) return "wheel";
  return "appear";
}

export function readMaster(
  view: DataView,
  references: Map<number, number>,
  group: SlideTextGroup,
  context: DrawingContext,
  name: string,
): MasterDefinition {
  const offset = references.get(group.persistReference);
  if (offset === undefined) {
    throw new PptParseError("Corrupt legacy PowerPoint file: missing master persist object");
  }
  const masterHeader = readRecordHeader(view, offset);
  const master = readRecordTree(view, offset, masterHeader.end)[0];
  if (!master || (master.type !== RecordType.slide && master.type !== RecordType.mainMaster)) {
    throw new PptParseError("Corrupt legacy PowerPoint file: unexpected master persist object");
  }
  const drawing = findDirect(master, RecordType.ppDrawing);
  const drawingChildren = drawing ? readDrawingChildren(view, drawing, context) : [];
  const children = appendUnreferencedText(drawingChildren, context);
  return children.length > 0 ? { name, children } : { name };
}

function appendUnreferencedText(children: SlideChild[], context: DrawingContext): SlideChild[] {
  const result = [...children];
  const projectedText = children.flatMap(slideChildText).map(compactText);
  for (const [index, entry] of context.entries.entries()) {
    if (context.usedEntryIndexes.has(index)) continue;
    if (projectedText.includes(compactText(entry.text))) continue;
    result.push({
      shape: {
        textBox: true,
        textBody: createTextBody([entry]),
      },
    });
  }
  return result;
}

function slideChildText(child: SlideChild): string[] {
  if ("shape" in child) return [textBodyToText(child.shape.textBody)];
  if ("group" in child) return child.group.children.flatMap(slideChildText);
  if ("table" in child) {
    return child.table.rows.flatMap((row) =>
      row.cells.map((cell) => (typeof cell.text === "string" ? cell.text : textBodyToText(cell))),
    );
  }
  return [];
}

function compactText(value: string): string {
  return value.replace(/\s+/g, "");
}
