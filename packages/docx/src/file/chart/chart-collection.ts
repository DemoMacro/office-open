/**
 * Chart collection module for managing chart data parts.
 *
 * @module
 */
import type { XmlComponent } from "@file/xml-components";

/**
 * Stores chart XML parts for later serialization by the compiler.
 */
export interface IChartData {
    /** Unique key for this chart (e.g., "chart1") */
    readonly key: string;
    /** The ChartSpace XML component */
    readonly chartSpace: XmlComponent;
}

/**
 * Manages chart parts in a document.
 *
 * Similar to Media, this collection stores chart XML components
 * that will be serialized into separate XML parts in the DOCX package.
 */
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
