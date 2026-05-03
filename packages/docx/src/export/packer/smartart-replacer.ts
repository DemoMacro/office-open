/**
 * SmartArt replacer module for substituting SmartArt placeholders in XML.
 *
 * @module
 */
import type { RelationshipType } from "@file/relationships/relationship/relationship";
import type { SmartArtCollection } from "@file/smartart/smartart-collection";

/**
 * Replaces SmartArt placeholder tokens with relationship IDs in XML content.
 *
 * SmartArt uses multiple placeholders:
 * - `{smartart:N}` — data model relationship (internal)
 * - `{smartart-lo:N}` — layout relationship (internal)
 * - `{smartart-qs:N}` — quick style relationship (internal)
 * - `{smartart-cs:N}` — color style relationship (internal)
 */
export class SmartArtReplacer {
    /**
     * Replaces SmartArt placeholder tokens with relationship IDs.
     */
    public replace(xmlData: string, smartArts: SmartArtCollection, dataOffset: number): string {
        let currentXmlData = xmlData;

        smartArts.Array.forEach((smartArtData, i) => {
            const key = smartArtData.key;

            // Data model: internal relationship to diagrams/data{n}.xml
            currentXmlData = currentXmlData.replace(
                new RegExp(`\\{smartart:${key}\\}`, "g"),
                (dataOffset + i).toString(),
            );

            // Layout, style, color: internal relationships to local files
            // Offset: dataOffset + smartArtCount for layout, +1 for style, +2 for color
            const loOffset = dataOffset + smartArts.Array.length;
            const qsOffset = loOffset + smartArts.Array.length;
            const csOffset = qsOffset + smartArts.Array.length;

            currentXmlData = currentXmlData.replace(
                new RegExp(`\\{smartart-lo:${key}\\}`, "g"),
                (loOffset + i).toString(),
            );
            currentXmlData = currentXmlData.replace(
                new RegExp(`\\{smartart-qs:${key}\\}`, "g"),
                (qsOffset + i).toString(),
            );
            currentXmlData = currentXmlData.replace(
                new RegExp(`\\{smartart-cs:${key}\\}`, "g"),
                (csOffset + i).toString(),
            );
        });

        return currentXmlData;
    }

    /**
     * Adds SmartArt relationships to the document relationships.
     *
     * All relationships are internal (pointing to package-local files).
     */
    public addRelationships(
        smartArts: SmartArtCollection,
        addRelationship: (
            id: number,
            type: RelationshipType,
            target: string,
            targetMode?: string,
        ) => void,
        baseOffset: number,
        chartCount: number,
    ): void {
        const dataOffset = baseOffset + chartCount;
        const smartArtCount = smartArts.Array.length;
        const loOffset = dataOffset + smartArtCount;
        const qsOffset = loOffset + smartArtCount;
        const csOffset = qsOffset + smartArtCount;

        smartArts.Array.forEach((_smartArtData, i) => {
            // Data model relationship (internal)
            addRelationship(
                dataOffset + i,
                "http://schemas.openxmlformats.org/officeDocument/2006/relationships/diagramData",
                `diagrams/data${i + 1}.xml`,
            );

            // Layout relationship (internal)
            addRelationship(
                loOffset + i,
                "http://schemas.microsoft.com/office/2007/relationships/diagramLayout",
                `diagrams/layout${i + 1}.xml`,
            );

            // Quick style relationship (internal)
            addRelationship(
                qsOffset + i,
                "http://schemas.microsoft.com/office/2007/relationships/diagramStyle",
                `diagrams/quickStyle${i + 1}.xml`,
            );

            // Color relationship (internal)
            addRelationship(
                csOffset + i,
                "http://schemas.microsoft.com/office/2007/relationships/diagramColors",
                `diagrams/colors${i + 1}.xml`,
            );

            // Drawing cache relationship (internal)
            const drOffset = csOffset + smartArtCount;
            addRelationship(
                drOffset + i,
                "http://schemas.microsoft.com/office/2007/relationships/diagramDrawing",
                `diagrams/drawing${i + 1}.xml`,
            );
        });
    }
}
