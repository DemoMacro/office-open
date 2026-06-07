export interface HyperlinkData {
  key: string;
  url: string;
  tooltip?: string;
}

export class HyperlinkCollection {
  private map = new Map<string, HyperlinkData>();

  public addHyperlink(key: string, url: string, tooltip?: string): void {
    this.map.set(key, { key, url, tooltip });
  }

  public get array(): HyperlinkData[] {
    return [...this.map.values()];
  }
}
