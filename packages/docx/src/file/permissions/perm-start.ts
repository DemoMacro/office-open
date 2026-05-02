/**
 * Permission range markers for WordprocessingML document protection.
 *
 * These elements mark editable ranges within a protected document.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, CT_PermStart, CT_PermEnd
 *
 * @module
 */
import { XmlAttributeComponent, XmlComponent } from "@file/xml-components";

/**
 * Editing group values for permission ranges.
 *
 * Defines who can edit within a permission range.
 *
 * Reference: ISO/IEC 29500-4, ST_EdGrp
 */
export const EditGroupType = {
    NONE: "none",
    EVERYONE: "everyone",
    ADMINISTRATORS: "administrators",
    CONTRIBUTORS: "contributors",
    EDITORS: "editors",
    OWNERS: "owners",
    CURRENT: "current",
} as const;

export type EditGroup = (typeof EditGroupType)[keyof typeof EditGroupType];

/**
 * Options for creating a permission start marker.
 */
export interface IPermStartOptions {
    /** Unique identifier for this permission range (typically a number) */
    readonly id: string | number;
    /** Editing group that can edit this range */
    readonly edGroup?: EditGroup;
    /** Individual user who can edit this range */
    readonly ed?: string;
    /** First column this range covers (for table cells) */
    readonly colFirst?: number;
    /** Last column this range covers (for table cells) */
    readonly colLast?: number;
}

/**
 * @internal
 */
class PermStartAttributes extends XmlAttributeComponent<IPermStartOptions> {
    protected readonly xmlKeys = {
        colFirst: "w:colFirst",
        colLast: "w:colLast",
        ed: "w:ed",
        edGroup: "w:edGrp",
        id: "w:id",
    };
}

/**
 * @internal
 */
class PermEndAttributes extends XmlAttributeComponent<{ readonly id: string | number }> {
    protected readonly xmlKeys = { id: "w:id" };
}

/**
 * Marks the start of a permission range.
 *
 * Permission ranges define editable regions within a protected document.
 * The range extends from this element to the matching `PermEnd`.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, permStart (CT_PermStart)
 *
 * @example
 * ```typescript
 * new PermStart({ id: "editRange1", edGroup: "everyone" });
 * ```
 */
export class PermStart extends XmlComponent {
    public constructor(options: IPermStartOptions) {
        super("w:permStart");
        this.root.push(new PermStartAttributes(options));
    }
}

/**
 * Marks the end of a permission range.
 *
 * Must be paired with a preceding `PermStart` with the same id.
 *
 * Reference: ISO/IEC 29500-4, wml.xsd, permEnd (CT_Perm)
 *
 * @example
 * ```typescript
 * new PermEnd("editRange1");
 * ```
 */
export class PermEnd extends XmlComponent {
    public constructor(id: string | number) {
        super("w:permEnd");
        this.root.push(new PermEndAttributes({ id }));
    }
}
