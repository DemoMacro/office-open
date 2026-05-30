import type { AnimationOptions } from "@file/animation/types";
import type { EffectsOptions } from "@file/drawingml/effects";
import type { OutlineOptions } from "@file/drawingml/outline";
import { ShapeProperties } from "@file/drawingml/shape-properties";
import type { ShapePropertiesOptions } from "@file/drawingml/shape-properties";
import type { File } from "@file/file";
import { XmlComponent as Xc } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import { escapeXml } from "@office-open/xml";
import { emuPositionOptional } from "@util/position";

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
  readonly textBody?: TextBodyOptions;
  readonly animation?: AnimationOptions;
  readonly placeholder?: "title" | "body" | "subTitle" | "sldNum" | "dt" | "ftr" | "hdr" | "obj";
  readonly placeholderIndex?: number;
}

/**
 * p:sp — A shape on a slide.
 * Lazy: stores options, builds XML object in prepForXml.
 *
 * x/y/width/height accept pixel values and are internally converted to EMUs.
 */
export class Shape extends Xc {
  private static nextId = 2;
  public readonly shapeId: number;
  public readonly animation: AnimationOptions | undefined;
  private readonly options: ShapeOptions;

  public constructor(options: ShapeOptions = {}) {
    super("p:sp");

    const id = options.id ?? Shape.nextId++;
    this.shapeId = id;
    this.animation = options.animation;
    this.options = { ...options, id };
  }

  public override toXml(context: Context): string {
    const opts = this.options;
    const id = this.shapeId;
    const name = opts.name ?? `Shape ${id}`;
    const parts: string[] = [];

    // p:nvSpPr
    let nvPrContent = "<p:nvPr/>";
    if (opts.placeholder) {
      const phAttrs: string[] = [`type="${opts.placeholder}"`];
      if (opts.placeholderIndex !== undefined) phAttrs.push(`idx="${opts.placeholderIndex}"`);
      nvPrContent = `<p:nvPr><p:ph ${phAttrs.join(" ")}/></p:nvPr>`;
    }
    parts.push(
      `<p:nvSpPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/><p:cNvSpPr/>${nvPrContent}</p:nvSpPr>`,
    );

    // p:spPr (ShapeProperties — has side effects, uses prepForXml → xml internally)
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
    const spPrXml = spPr.toXml(context as Context<File>);
    if (spPrXml) parts.push(spPrXml);

    // p:txBody (TextBody — has toXml)
    const txBody = new TextBody(opts.textBody ?? {});
    parts.push(txBody.toXml(context));

    return `<p:sp>${parts.join("")}</p:sp>`;
  }
}
