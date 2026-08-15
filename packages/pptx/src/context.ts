/**
 * PPTX write context — collects resources during descriptor-driven serialization.
 *
 * @module
 */

import { EmbeddingCollection, Media } from "@office-open/core";
import type { BaseMediaEntry, EmbeddingData } from "@office-open/core";
import type { HyperlinkTarget, ReadContext, WriteContext } from "@office-open/core/descriptor";
import type {
  ColorDefinitionOptions,
  LayoutDefinitionOptions,
  StyleDefinitionOptions,
} from "@office-open/core/smartart";

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

export interface MediaEntry extends BaseMediaEntry {
  key: string;
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
  /** Built-in layout id or a full custom layout definition. */
  layout: string | LayoutDefinitionOptions;
  /** Built-in quick-style id or a full custom style definition. */
  style: string | StyleDefinitionOptions;
  /** Built-in color-transform id or a full custom color definition. */
  color: string | ColorDefinitionOptions;
}

export interface HyperlinkEntry {
  key: string;
  url?: string;
  slide?: number;
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
  private _media = new Media<MediaEntry>();
  private _embeddings = new EmbeddingCollection();
  private _charts = new Map<string, ChartEntry>();
  private _smartArts = new Map<string, SmartArtEntry>();
  private _hyperlinks = new Map<string, HyperlinkEntry>();
  private _nextRelId = 1;
  private _nextChartId = 1;
  private _nextSmartArtId = 1;

  /**
   * Slide width in EMU — the master's standard placeholder positions are scaled
   * to this width. Set by the compiler from the presentation size; defaults to
   * the 16:9 width (12192000 EMU) so patch/deserialize paths that never touch
   * the master still have a sane value.
   */
  public slideWidth = 12192000;

  // ── WriteContext stubs (core interface) ──

  public addRelationship(_type: string, _target: string, _mode?: string): string {
    const id = this._nextRelId++;
    return `rId${id}`;
  }

  public addMedia(data: Uint8Array, type: string): string {
    const entry = this._media.addMedia(data, type, (fileName) => ({
      key: fileName,
      fileName,
      data,
      type,
      transformation: { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } },
    }));
    return `{${entry.fileName}}`;
  }

  // ── PPTX-specific registration ──

  /**
   * Register an OLE embedding (ppt/embeddings/oleObjectN.bin) and return a
   * `{ole:oleObjectN.bin}` placeholder. The compiler rewrites the placeholder
   * to a real relationship id and adds the oleObject relationship per slide.
   */
  public addOle(data: Uint8Array, progId?: string): string {
    const fileName = this._embeddings.nextEmbeddingName();
    const entry: EmbeddingData = {
      fileName,
      data,
      ...(progId !== undefined ? { progId } : {}),
    };
    this._embeddings.addEmbedding(fileName, entry);
    return `{ole:${fileName}}`;
  }

  public addImage(key: string, entry: MediaEntry): MediaEntry {
    return this._media.addMedia(
      entry.data,
      entry.type,
      (fileName) => ({ ...entry, fileName, key: fileName }),
      key,
    );
  }

  public addChart(key: string, entry: ChartEntry): void {
    this._charts.set(key, entry);
  }

  public addSmartArt(key: string, entry: SmartArtEntry): void {
    this._smartArts.set(key, entry);
  }

  public addHyperlink(key: string, target: HyperlinkTarget): void {
    this._hyperlinks.set(key, {
      key,
      url: target.url,
      slide: target.slide,
      tooltip: target.tooltip,
    });
  }

  public nextChartKey(): string {
    return `chart_${this._nextChartId++}`;
  }

  public nextSmartArtKey(): string {
    return `smartart_${this._nextSmartArtId++}`;
  }

  // ── Getters ──

  public get media(): MediaEntry[] {
    return this._media.array;
  }

  /** Underlying deduplicated collection — used by the compiler for media output. */
  public get mediaCollection(): Media<MediaEntry> {
    return this._media;
  }

  /** Registered OLE embeddings — output as ppt/embeddings/*.bin by the compiler. */
  public get embeddings(): EmbeddingData[] {
    return this._embeddings.array;
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
