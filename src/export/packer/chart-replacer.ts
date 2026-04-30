/**
 * Chart replacer module for substituting chart placeholders in XML.
 *
 * @module
 */
import type { ChartCollection } from "@file/chart/chart-collection";

/**
 * Replaces chart placeholder tokens with relationship IDs in XML content.
 *
 * Charts use placeholders like `{chart:chart_123}` in the document XML.
 * This class replaces them with the actual relationship IDs.
 */
export class ChartReplacer {
    /**
     * Replaces chart placeholder tokens with relationship IDs.
     *
     * @param xmlData - The XML string containing chart placeholders
     * @param charts - The chart collection
     * @param offset - Starting offset for relationship IDs
     * @returns XML string with placeholders replaced by relationship IDs
     */
    public replace(xmlData: string, charts: ChartCollection, offset: number): string {
        let currentXmlData = xmlData;

        charts.Array.forEach((chartData, i) => {
            currentXmlData = currentXmlData.replace(
                new RegExp(`\\{chart:${chartData.key}\\}`, "g"),
                (offset + i).toString(),
            );
        });

        return currentXmlData;
    }
}
