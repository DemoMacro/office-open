export interface IHyperlinkData {
    readonly key: string;
    readonly url: string;
    readonly tooltip?: string;
}

export class HyperlinkCollection {
    private readonly map = new Map<string, IHyperlinkData>();

    public addHyperlink(key: string, url: string, tooltip?: string): void {
        this.map.set(key, { key, url, tooltip });
    }

    public get Array(): readonly IHyperlinkData[] {
        return [...this.map.values()];
    }
}
