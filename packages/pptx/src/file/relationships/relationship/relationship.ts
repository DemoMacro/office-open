import { BuilderElement } from "@file/xml-components";
import type { XmlComponent } from "@file/xml-components";

export type RelationshipType =
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    | "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"
    | "http://schemas.microsoft.com/office/2007/relationships/diagramLayout"
    | "http://schemas.microsoft.com/office/2007/relationships/diagramStyle"
    | "http://schemas.microsoft.com/office/2007/relationships/diagramColors";

export const TargetModeType = {
    EXTERNAL: "External",
} as const;

interface IRelationshipAttributes {
    readonly id: string;
    readonly type: RelationshipType;
    readonly target: string;
    readonly targetMode?: (typeof TargetModeType)[keyof typeof TargetModeType];
}

export const createRelationship = (
    id: string,
    type: RelationshipType,
    target: string,
    targetMode?: (typeof TargetModeType)[keyof typeof TargetModeType],
): XmlComponent =>
    new BuilderElement<IRelationshipAttributes>({
        attributes: {
            id: { key: "Id", value: id },
            target: { key: "Target", value: target },
            targetMode: { key: "TargetMode", value: targetMode },
            type: { key: "Type", value: type },
        },
        name: "Relationship",
    });
