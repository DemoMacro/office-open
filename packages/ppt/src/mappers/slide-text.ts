import {
  RecordType,
  decodeAnsi,
  decodeUtf16,
  readInt32,
  readUint32,
  recordBody,
  type RecordNode,
} from "../records";
import type { SlideTextGroup } from "./models";
export function readSlideTextGroups(
  view: DataView,
  document: RecordNode,
  kind: "slide" | "master" | "notes",
): SlideTextGroup[] {
  const lists = document.children
    .filter(
      (child) =>
        child.type === RecordType.slideListWithText &&
        child.instance === (kind === "slide" ? 0 : kind === "master" ? 1 : 2),
    )
    .flatMap((list) => groupSlideText(view, list));
  return lists;
}

function groupSlideText(view: DataView, list: RecordNode): SlideTextGroup[] {
  const groups: SlideTextGroup[] = [];
  let current: SlideTextGroup | undefined;
  let placeholderType: number | undefined;

  for (const child of list.children) {
    if (child.type === RecordType.slidePersistAtom) {
      if (current) groups.push(current);
      current = {
        persistReference: readInt32(view, child, 0),
        slideIdentifier: child.length >= 16 ? readUint32(view, child, 12) : undefined,
        entries: [],
      };
      placeholderType = undefined;
      continue;
    }
    if (child.type === RecordType.textHeader) {
      placeholderType = readInt32(view, child, 0);
      continue;
    }
    if (current && placeholderType !== undefined) {
      const bytes =
        child.type === RecordType.textChars || child.type === RecordType.textBytes
          ? recordBody(view, child)
          : undefined;
      if (bytes) {
        current.entries.push({
          placeholderType,
          text: child.type === RecordType.textChars ? decodeUtf16(bytes) : decodeAnsi(bytes),
        });
        placeholderType = undefined;
      }
    }
  }
  if (current) groups.push(current);
  return groups;
}
