/**
 * PPTX write context — collects resources during descriptor-driven serialization.
 *
 * @module
 */

import type { ReadContext, WriteContext } from "@office-open/core/descriptor";

import type { ParseContext } from "../parse/context";

// ── Resource entries ──

export interface MediaEntry {
  readonly key: string;
  readonly fileName: string;
  readonly data: Uint8Array;
  readonly type: string;
  readonly transformation: {
    readonly pixels: { readonly x: number; readonly y: number };
    readonly emus: { readonly x: number; readonly y: number };
  };
}

export interface ChartEntry {
  readonly key: string;
  readonly chartSpaceXml: string;
}

export interface SmartArtEntry {
  readonly key: string;
  readonly dataModelXml: string;
  readonly layout: string;
  readonly style: string;
  readonly color: string;
}

export interface HyperlinkEntry {
  readonly key: string;
  readonly url: string;
  readonly tooltip?: string;
}

// ── Context ──

/**
 * PPTX-specific write context.
 *
 * Extends the core {@link WriteContext} with media, chart, SmartArt,
 * and hyperlink registration methods. Custom descriptors cast the
 * generic `WriteContext` to this type to access PPTX-specific features.
 */
export class PptxWriteContext implements WriteContext {
  private readonly _media = new Map<string, MediaEntry>();
  private readonly _charts = new Map<string, ChartEntry>();
  private readonly _smartArts = new Map<string, SmartArtEntry>();
  private readonly _hyperlinks = new Map<string, HyperlinkEntry>();
  private readonly _contentTypes = new Map<string, string>();
  private _nextRelId = 1;
  private _nextMediaId = 1;
  private _nextChartId = 1;
  private _nextSmartArtId = 1;

  // ── WriteContext stubs (core interface) ──

  public addRelationship(_type: string, _target: string, _mode?: string): string {
    const id = this._nextRelId++;
    return `rId${id}`;
  }

  public addMedia(_data: Uint8Array, _type: string): string {
    return `media_${this._nextMediaId++}`;
  }

  // ── PPTX-specific registration ──

  public addImage(key: string, entry: MediaEntry): void {
    this._media.set(key, entry);
  }

  public addChart(key: string, entry: ChartEntry): void {
    this._charts.set(key, entry);
  }

  public addSmartArt(key: string, entry: SmartArtEntry): void {
    this._smartArts.set(key, entry);
  }

  public addHyperlink(key: string, url: string, tooltip?: string): void {
    this._hyperlinks.set(key, { key, url, tooltip });
  }

  public addContentType(extension: string, contentType: string): void {
    this._contentTypes.set(extension, contentType);
  }

  public nextChartKey(): string {
    return `chart_${this._nextChartId++}`;
  }

  public nextSmartArtKey(): string {
    return `smartart_${this._nextSmartArtId++}`;
  }

  // ── Getters ──

  public get media(): readonly MediaEntry[] {
    return [...this._media.values()];
  }

  public get charts(): readonly ChartEntry[] {
    return [...this._charts.values()];
  }

  public get smartArts(): readonly SmartArtEntry[] {
    return [...this._smartArts.values()];
  }

  public get hyperlinks(): readonly HyperlinkEntry[] {
    return [...this._hyperlinks.values()];
  }

  public get contentTypes(): ReadonlyMap<string, string> {
    return this._contentTypes;
  }
}

// ── Read context ──

/**
 * PPTX-specific read context.
 *
 * Adapts the existing {@link ParseContext} to the core {@link ReadContext}
 * interface used by the descriptor parse pipeline.
 */
export class PptxReadContext implements ReadContext {
  constructor(private readonly _parseCtx: ParseContext) {}

  public resolveRelationship(rId: string): string | undefined {
    return this._parseCtx.slideRels.get(rId);
  }

  public getPart(path: string) {
    return this._parseCtx.pptx.doc.get(path);
  }

  public getRaw(path: string): Uint8Array | undefined {
    return this._parseCtx.pptx.doc.getRaw(path);
  }
}
