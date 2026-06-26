/**
 * PPTX write context — collects resources during descriptor-driven serialization.
 *
 * @module
 */

import type { ReadContext, WriteContext } from "@office-open/core/descriptor";

import type { PptxDocument } from "./parse";

// ── Parse context (moved from parse/context.ts) ──

/**
 * Parse context for PPTX documents.
 */
export class ParseContext {
  constructor(
    public pptx: PptxDocument,
    /** Slide relationship ID → path, parsed from slide's _rels file */
    public slideRels: Map<string, string>,
  ) {}
}

// ── Resource entries ──

export interface MediaEntry {
  key: string;
  fileName: string;
  data: Uint8Array;
  type: string;
  transformation: {
    pixels: { x: number; y: number };
    emus: { x: number; y: number };
  };
}

export interface ChartEntry {
  key: string;
  chartSpaceXml: string;
}

export interface SmartArtEntry {
  key: string;
  dataModelXml: string;
  layout: string;
  style: string;
  color: string;
}

export interface HyperlinkEntry {
  key: string;
  url: string;
  tooltip?: string;
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
  private _media = new Map<string, MediaEntry>();
  private _charts = new Map<string, ChartEntry>();
  private _smartArts = new Map<string, SmartArtEntry>();
  private _hyperlinks = new Map<string, HyperlinkEntry>();
  private _contentTypes = new Map<string, string>();
  private _nextRelId = 1;
  private _nextMediaId = 1;
  private _nextChartId = 1;
  private _nextSmartArtId = 1;

  // ── WriteContext stubs (core interface) ──

  public addRelationship(_type: string, _target: string, _mode?: string): string {
    const id = this._nextRelId++;
    return `rId${id}`;
  }

  public addMedia(data: Uint8Array, type: string): string {
    const fileName = `image${this._nextMediaId++}.${type}`;
    this._media.set(fileName, {
      key: fileName,
      fileName,
      data,
      type,
      transformation: { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } },
    });
    return `{${fileName}}`;
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

  public get media(): MediaEntry[] {
    return [...this._media.values()];
  }

  public get charts(): ChartEntry[] {
    return [...this._charts.values()];
  }

  public get smartArts(): SmartArtEntry[] {
    return [...this._smartArts.values()];
  }

  public get hyperlinks(): HyperlinkEntry[] {
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
  constructor(private _parseCtx: ParseContext) {}

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
