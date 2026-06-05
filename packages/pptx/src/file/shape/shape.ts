import type { AnimationOptions } from "@file/animation/types";
import type { EffectsOptions } from "@file/drawingml/effects";
import type { OutlineOptions } from "@file/drawingml/outline";
import { ShapeProperties } from "@file/drawingml/shape-properties";
import type { ShapePropertiesOptions } from "@file/drawingml/shape-properties";
import type { File } from "@file/file";
import { XmlComponent as Xc } from "@file/xml-components";
import type { Context } from "@file/xml-components";
import type { ShapeLockingOptions } from "@office-open/core";
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
  readonly locking?: ShapeLockingOptions;
  readonly placeholder?: "title" | "body" | "subTitle" | "sldNum" | "dt" | "ftr" | "hdr" | "obj";
  readonly placeholderIndex?: number;
  readonly useBackgroundFill?: boolean;
  readonly isPhoto?: boolean;
  readonly userDrawn?: boolean;
  readonly hasCustomPrompt?: boolean;
  /** Black-and-white mode for the shape. */
  readonly blackWhiteMode?:
    | "clr"
    | "auto"
    | "gray"
    | "ltGray"
    | "invGray"
    | "grayWhite"
    | "blackGray"
    | "blackWhite"
    | "black"
    | "white"
    | "hidden";
}

/**
 * p:sp — A shape on a slide.
 * Lazy: stores options, builds XML in toXml().
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
      if (opts.hasCustomPrompt) phAttrs.push('hasCustomPrompt="1"');
      nvPrContent = `<p:nvPr><p:ph ${phAttrs.join(" ")}/></p:nvPr>`;
    } else if (opts.isPhoto || opts.userDrawn) {
      const nvPrAttrs: string[] = [];
      if (opts.isPhoto) nvPrAttrs.push('isPhoto="1"');
      if (opts.userDrawn) nvPrAttrs.push('userDrawn="1"');
      nvPrContent = `<p:nvPr ${nvPrAttrs.join(" ")}/>`;
    }
    // a:spLocks inside p:cNvSpPr
    let cNvSpPrContent = "<p:cNvSpPr/>";
    if (opts.locking) {
      const lockAttrs = buildLockAttrs(opts.locking);
      if (lockAttrs.length > 0) {
        cNvSpPrContent = `<p:cNvSpPr><a:spLocks ${lockAttrs.join(" ")}/></p:cNvSpPr>`;
      }
    }
    parts.push(
      `<p:nvSpPr><p:cNvPr id="${id}" name="${escapeXml(name)}"/>${cNvSpPrContent}${nvPrContent}</p:nvSpPr>`,
    );

    // p:spPr (ShapeProperties — has side effects, uses toXml() internally)
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

    const spAttrs: string[] = [];
    if (opts.useBackgroundFill) spAttrs.push(' useBgFill="1"');
    if (opts.blackWhiteMode) spAttrs.push(` bwMode="${opts.blackWhiteMode}"`);
    return `<p:sp${spAttrs.join("")}>${parts.join("")}</p:sp>`;
  }
}

/** Build locking attribute string[] from ShapeLockingOptions. */
function buildLockAttrs(opts: ShapeLockingOptions): string[] {
  const attrs: string[] = [];
  const keys = [
    "noGrp",
    "noSelect",
    "noRot",
    "noChangeAspect",
    "noMove",
    "noResize",
    "noEditPoints",
    "noAdjustHandles",
    "noChangeArrowheads",
    "noChangeShapeType",
    "noTextEdit",
  ] as const;
  for (const key of keys) {
    const val = opts[key];
    if (val !== undefined) attrs.push(`${key}="${val ? 1 : 0}"`);
  }
  return attrs;
}
