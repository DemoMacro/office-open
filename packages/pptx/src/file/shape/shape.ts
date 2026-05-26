import type { AnimationOptions } from "@file/animation/types";
import type { EffectsOptions } from "@file/drawingml/effects";
import type { OutlineOptions } from "@file/drawingml/outline";
import { ShapeProperties } from "@file/drawingml/shape-properties";
import type { ShapePropertiesOptions } from "@file/drawingml/shape-properties";
import type { File } from "@file/file";
import { XmlComponent as Xc } from "@file/xml-components";
import type { Context, IXmlableObject } from "@file/xml-components";
import { emuPositionOptional } from "@util/position";

import { Paragraph } from "./paragraph/paragraph";
import { TextBody } from "./text-body";
import type { TextBodyOptions } from "./text-body";

export interface ShapeOptions {
  readonly id?: number;
  readonly name?: string;
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly geometry?: string;
  readonly fill?: ShapePropertiesOptions["fill"];
  readonly outline?: OutlineOptions;
  readonly effects?: EffectsOptions;
  readonly flipHorizontal?: boolean;
  readonly rotation?: number;
  readonly text?: string;
  readonly paragraphs?: TextBodyOptions["paragraphs"];
  readonly textVertical?: TextBodyOptions["vertical"];
  readonly textAnchor?: TextBodyOptions["anchor"];
  readonly textAutoFit?: TextBodyOptions["autoFit"];
  readonly textWrap?: TextBodyOptions["wrap"];
  readonly textMargins?: TextBodyOptions["margins"];
  readonly textColumns?: TextBodyOptions["columns"];
  readonly textColumnSpacing?: TextBodyOptions["columnSpacing"];
  readonly animation?: AnimationOptions;
  readonly placeholder?: "title" | "body" | "subTitle" | "sldNum" | "dt" | "ftr" | "hdr" | "obj";
  readonly placeholderIndex?: number;
}

/**
 * Pure function: builds p:ph element for placeholder.
 */
function buildPlaceholder(type: string, index?: number): IXmlableObject {
  const attrs: Record<string, string | number> = { type };
  if (index !== undefined) attrs.idx = index;
  return { "p:ph": { _attr: attrs } };
}

/**
 * p:sp — A shape on a slide.
 * Lazy: stores options, builds XML object in prepForXml.
 *
 * x/y/width/height accept pixel values and are internally converted to EMUs.
 */
export class Shape extends Xc {
  private static nextId = 2;
  private readonly shapeId: number;
  private readonly animationOptions?: AnimationOptions;
  private readonly options: ShapeOptions;

  public constructor(options: ShapeOptions = {}) {
    super("p:sp");

    const id = options.id ?? Shape.nextId++;
    this.shapeId = id;
    this.animationOptions = options.animation;
    this.options = { ...options, id };
  }

  public get ShapeId(): number {
    return this.shapeId;
  }

  public get Animation(): AnimationOptions | undefined {
    return this.animationOptions;
  }

  public override prepForXml(context: Context): IXmlableObject | undefined {
    const opts = this.options;
    const id = this.shapeId;
    const name = opts.name ?? `Shape ${id}`;
    const children: IXmlableObject[] = [];

    // nvSpPr
    const nvPrChildren: IXmlableObject[] = [];
    if (opts.placeholder) {
      nvPrChildren.push(buildPlaceholder(opts.placeholder, opts.placeholderIndex));
    }
    children.push({
      "p:nvSpPr": [
        { "p:cNvPr": { _attr: { id, name } } },
        { "p:cNvSpPr": {} },
        { "p:nvPr": nvPrChildren.length > 0 ? nvPrChildren : {} },
      ],
    });

    // spPr (ShapeProperties)
    const shapeProps: ShapePropertiesOptions = {
      ...emuPositionOptional(opts),
      geometry: opts.geometry,
      fill: opts.fill,
      outline: opts.outline,
      effects: opts.effects,
      flipHorizontal: opts.flipHorizontal,
      rotation: opts.rotation,
    };
    const spPr = new ShapeProperties(shapeProps);
    const spPrObj = spPr.prepForXml(context as Context<File>);
    if (spPrObj) children.push(spPrObj);

    // txBody (TextBody)
    const textBodyOptions: TextBodyOptions = {
      paragraphs: opts.paragraphs ?? (opts.text ? [new Paragraph({ text: opts.text })] : undefined),
      vertical: opts.textVertical,
      anchor: opts.textAnchor,
      autoFit: opts.textAutoFit,
      wrap: opts.textWrap,
      margins: opts.textMargins,
      columns: opts.textColumns,
      columnSpacing: opts.textColumnSpacing,
    };
    const txBody = new TextBody(textBodyOptions);
    const txBodyObj = txBody.prepForXml(context);
    if (txBodyObj) children.push(txBodyObj);

    return { "p:sp": children };
  }
}
