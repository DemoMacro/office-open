import { XmlComponent, XmlAttributeComponent, BuilderElement } from "../xml-components";

export type RelationshipType =
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
    | "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors"
    | "http://schemas.microsoft.com/office/2007/relationships/diagramLayout"
    | "http://schemas.microsoft.com/office/2007/relationships/diagramStyle"
    | "http://schemas.microsoft.com/office/2007/relationships/diagramColors"
    | "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing"
    // WordprocessingML specific
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/fontTable"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/webSettings"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/numbering"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/header"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/custom-properties"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/footnotes"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/endnotes"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/bibliography"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/font"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/subDocument"
    // PresentationML specific
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideMaster"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/slideLayout"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesSlide"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/presProps"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/viewProps"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/tableStyles"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/notesMaster"
    | "http://schemas.openxmlformats.org/officeDocument/2006/relationships/video"
    | "http://schemas.microsoft.com/office/2007/relationships/media";

export const TargetModeType = {
    EXTERNAL: "External",
} as const;

class RelationshipsAttributes extends XmlAttributeComponent<{
    readonly xmlns: string;
}> {
    protected readonly xmlKeys = {
        xmlns: "xmlns",
    };
}

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

export class Relationships extends XmlComponent {
    public constructor() {
        super("Relationships");
        this.root.push(
            new RelationshipsAttributes({
                xmlns: "http://schemas.openxmlformats.org/package/2006/relationships",
            }),
        );
    }

    public addRelationship(
        id: number | string,
        type: RelationshipType,
        target: string,
        targetMode?: (typeof TargetModeType)[keyof typeof TargetModeType],
    ): void {
        this.root.push(createRelationship(`rId${id}`, type, target, targetMode));
    }

    public get RelationshipCount(): number {
        return this.root.length - 1;
    }
}
