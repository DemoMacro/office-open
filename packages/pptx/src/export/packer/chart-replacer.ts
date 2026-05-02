import type { ChartCollection } from "@file/chart/chart-collection";

export class ChartReplacer {
    public replace(xmlData: string, charts: ChartCollection, offset: number): string {
        let currentXmlData = xmlData;

        charts.Array.forEach((chartData, i) => {
            currentXmlData = currentXmlData.replace(
                new RegExp(`\\{chart:${chartData.key}\\}`, "g"),
                `rId${offset + i}`,
            );
        });

        return currentXmlData;
    }
}
