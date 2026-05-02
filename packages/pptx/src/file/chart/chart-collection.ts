import type { XmlComponent } from "@file/xml-components";

export interface IChartData {
    readonly key: string;
    readonly chartSpace: XmlComponent;
}

export class ChartCollection {
    private readonly map: Map<string, IChartData>;

    public constructor() {
        this.map = new Map<string, IChartData>();
    }

    public addChart(key: string, chartData: IChartData): void {
        this.map.set(key, chartData);
    }

    public get Array(): readonly IChartData[] {
        return [...this.map.values()];
    }
}
