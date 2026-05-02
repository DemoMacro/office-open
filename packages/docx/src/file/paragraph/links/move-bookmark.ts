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
import { XmlAttributeComponent, XmlComponent } from "@file/xml-components";

/**
 * @internal
 */
class MoveBookmarkAttributes extends XmlAttributeComponent<{
    readonly id: number;
    readonly name?: string;
    readonly author?: string;
    readonly date?: string;
}> {
    protected readonly xmlKeys = {
        author: "w:author",
        date: "w:date",
        id: "w:id",
        name: "w:name",
    };
}

/**
 * @internal
 */
class MoveRangeEndAttributes extends XmlAttributeComponent<{ readonly id: number }> {
    protected readonly xmlKeys = { id: "w:id" };
}

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
        this.root.push(new MoveBookmarkAttributes({ id, name, author, date }));
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
        this.root.push(new MoveRangeEndAttributes({ id }));
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
        this.root.push(new MoveBookmarkAttributes({ id, name, author, date }));
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
        this.root.push(new MoveRangeEndAttributes({ id }));
    }
}
