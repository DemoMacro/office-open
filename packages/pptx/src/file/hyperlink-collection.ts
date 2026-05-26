export interface HyperlinkData {
  readonly key: string;
  readonly url: string;
  readonly tooltip?: string;
}

export class HyperlinkCollection {
  private readonly map = new Map<string, HyperlinkData>();

  public addHyperlink(key: string, url: string, tooltip?: string): void {
    this.map.set(key, { key, url, tooltip });
  }

  public get Array(): readonly HyperlinkData[] {
    return [...this.map.values()];
  }
}
