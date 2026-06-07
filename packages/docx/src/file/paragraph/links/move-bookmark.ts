/**
 * Move bookmark range markers for WordprocessingML track changes.
 *
 * These elements mark the source and destination ranges of content moves
 * in track-revised documents.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_MoveBookmark
 *
 * @module
 */
import { XmlComponent } from "@file/xml-components";

/** Builds move bookmark attributes with optional fields omitted when undefined. */
const buildMoveBookmarkAttr = (
  id: number,
  name?: string,
  author?: string,
  date?: string,
): { _attr: Record<string, string | number> } => {
  const attrs: Record<string, string | number> = { "w:id": id };
  if (name !== undefined) attrs["w:name"] = name;
  if (author !== undefined) attrs["w:author"] = author;
  if (date !== undefined) attrs["w:date"] = date;
  return { _attr: attrs };
};

/**
 * Marks the start of a move source range.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, moveFromRangeStart (CT_MoveBookmark)
 *
 * @example
 * ```typescript
 * new MoveFromRangeStart(0, "movedPara", "John", "2024-01-01T00:00:00Z");
 * ```
 */
export class MoveFromRangeStart extends XmlComponent {
  public constructor(id: number, name?: string, author?: string, date?: string) {
    super("w:moveFromRangeStart");
    this.root.push(buildMoveBookmarkAttr(id, name, author, date));
  }
}

/**
 * Marks the end of a move source range.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, moveFromRangeEnd (CT_MarkupRange)
 */
export class MoveFromRangeEnd extends XmlComponent {
  public constructor(id: number) {
    super("w:moveFromRangeEnd");
    this.root.push({ _attr: { "w:id": id } });
  }
}

/**
 * Marks the start of a move destination range.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, moveToRangeStart (CT_MoveBookmark)
 *
 * @example
 * ```typescript
 * new MoveToRangeStart(0, "movedPara", "John", "2024-01-01T00:00:00Z");
 * ```
 */
export class MoveToRangeStart extends XmlComponent {
  public constructor(id: number, name?: string, author?: string, date?: string) {
    super("w:moveToRangeStart");
    this.root.push(buildMoveBookmarkAttr(id, name, author, date));
  }
}

/**
 * Marks the end of a move destination range.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, moveToRangeEnd (CT_MarkupRange)
 */
export class MoveToRangeEnd extends XmlComponent {
  public constructor(id: number) {
    super("w:moveToRangeEnd");
    this.root.push({ _attr: { "w:id": id } });
  }
}
