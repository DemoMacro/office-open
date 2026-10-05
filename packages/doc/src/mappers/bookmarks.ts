import type { SectionChild } from "@office-open/docx";

import type { TextCharacter } from "./models";
import { paragraphIndexForCp } from "./paragraph-projection";

export function applyBookmarks(
  children: SectionChild[],
  bookmarks: readonly { id: number; name: string; start: number; end: number }[],
  characters: readonly TextCharacter[],
): SectionChild[] {
  const boundaries = new Map<number, Array<{ marker: SectionChild; before: boolean }>>();
  for (const bookmark of bookmarks) {
    const startIndex = paragraphIndexForCp(characters, bookmark.start);
    boundaries.set(startIndex, [
      ...(boundaries.get(startIndex) ?? []),
      { marker: { bookmarkStart: { id: bookmark.id, name: bookmark.name } }, before: true },
    ]);
    const endIndex = paragraphIndexForCp(characters, bookmark.end);
    boundaries.set(endIndex, [
      ...(boundaries.get(endIndex) ?? []),
      { marker: { bookmarkEnd: { id: bookmark.id } }, before: false },
    ]);
  }
  if (boundaries.size === 0) return children;
  return children.flatMap((child, index) => {
    const markers = boundaries.get(index) ?? [];
    return [
      ...markers.filter((item) => item.before).map((item) => item.marker),
      child,
      ...markers.filter((item) => !item.before).map((item) => item.marker),
    ];
  });
}
