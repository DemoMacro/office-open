import type { SmartArtCollection } from "@file/smartart/smartart-collection";
import type { RelationshipType } from "@office-open/core";

export class SmartArtReplacer {
    public replace(xmlData: string, smartArts: SmartArtCollection, dataOffset: number): string {
        let currentXmlData = xmlData;

        smartArts.Array.forEach((smartArtData, i) => {
            const key = smartArtData.key;

            const loOffset = dataOffset + smartArts.Array.length;
            const qsOffset = loOffset + smartArts.Array.length;
            const csOffset = qsOffset + smartArts.Array.length;

            currentXmlData = currentXmlData.replace(
                new RegExp(`\\{smartart:${key}\\}`, "g"),
                `rId${dataOffset + i}`,
            );
            currentXmlData = currentXmlData.replace(
                new RegExp(`\\{smartart-lo:${key}\\}`, "g"),
                `rId${loOffset + i}`,
            );
            currentXmlData = currentXmlData.replace(
                new RegExp(`\\{smartart-qs:${key}\\}`, "g"),
                `rId${qsOffset + i}`,
            );
            currentXmlData = currentXmlData.replace(
                new RegExp(`\\{smartart-cs:${key}\\}`, "g"),
                `rId${csOffset + i}`,
            );
        });

        return currentXmlData;
    }

    public addRelationships(
        smartArts: SmartArtCollection,
        addRelationship: (id: number, type: RelationshipType, target: string) => void,
        baseOffset: number,
        globalStartIndex: number,
    ): void {
        const dataOffset = baseOffset;
        const smartArtCount = smartArts.Array.length;
        const loOffset = dataOffset + smartArtCount;
        const qsOffset = loOffset + smartArtCount;
        const csOffset = qsOffset + smartArtCount;
        smartArts.Array.forEach((_smartArtData, i) => {
            const gi = globalStartIndex + i;
            addRelationship(
                dataOffset + i,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
                `../diagrams/data${gi + 1}.xml`,
            );
            addRelationship(
                loOffset + i,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramLayout",
                `../diagrams/layout${gi + 1}.xml`,
            );
            addRelationship(
                qsOffset + i,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramQuickStyle",
                `../diagrams/quickStyle${gi + 1}.xml`,
            );
            addRelationship(
                csOffset + i,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramColors",
                `../diagrams/colors${gi + 1}.xml`,
            );
            const drOffset = csOffset + smartArtCount;
            addRelationship(
                drOffset + i,
                "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing",
                `../diagrams/drawing${gi + 1}.xml`,
            );
        });
    }
}
