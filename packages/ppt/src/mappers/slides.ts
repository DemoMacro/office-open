import type { AnimationEntry, MasterDefinition, SlideChild, SlideOptions } from "@office-open/pptx";

import { readPersistedRecord } from "../parts/persist";
import {
  LegacyPowerPointError,
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
  notesGroups: readonly SlideTextGroup[],
): SlideOptions {
  const offset = references.get(group.persistReference);
  if (offset === undefined) {
    throw new LegacyPowerPointError("Corrupt legacy PowerPoint file: missing slide persist object");
  }
  const slideHeader = readRecordHeader(view, offset);
  const slide = readRecordTree(view, offset, slideHeader.end)[0];
  if (!slide || slide.type !== RecordType.slide) {
    throw new LegacyPowerPointError(
      "Corrupt legacy PowerPoint file: unexpected slide persist object",
    );
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
  const notes = readNotes(view, references, notesGroups, notesId);
  if (notes) result.notes = notes;
  const animations = drawing
    ? readSlideAnimations(view, drawing).map((animation) => projectAnimation(view, animation))
    : [];
  if (animations.length > 0) result.animations = animations;
  return result;
}

export function readNotes(
  view: DataView,
  references: Map<number, number>,
  notesGroups: readonly SlideTextGroup[],
  notesId: number | undefined,
): string | undefined {
  if (!notesId) return undefined;
  const group = notesGroups.find((entry) => entry.slideIdentifier === notesId);
  const offset = group && references.get(group.persistReference);
  if (offset === undefined) return undefined;
  try {
    const notes = readPersistedRecord(view, offset, [RecordType.notes]);
    const drawing = findDirect(notes, RecordType.ppDrawing);
    if (!drawing) return undefined;
    const paragraphs = collectEmbeddedText(view, drawing)
      .flatMap((value) => splitParagraphs(value))
      .filter((value) => value.length > 0);
    return paragraphs.length > 0 ? paragraphs.join("\n") : undefined;
  } catch {
    return undefined;
  }
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
    throw new LegacyPowerPointError(
      "Corrupt legacy PowerPoint file: missing master persist object",
    );
  }
  const masterHeader = readRecordHeader(view, offset);
  const master = readRecordTree(view, offset, masterHeader.end)[0];
  if (!master || (master.type !== RecordType.slide && master.type !== RecordType.mainMaster)) {
    throw new LegacyPowerPointError(
      "Corrupt legacy PowerPoint file: unexpected master persist object",
    );
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
        textBody: createTextBody([entry.text]),
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
