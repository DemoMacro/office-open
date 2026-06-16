/**
 * Blip fill module for DrawingML pictures.
 *
 * This module defines how an image (blip) fills a picture shape,
 * including stretching and cropping options.
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_BlipFillProperties
 *
 * @module
 */
import { element } from "@office-open/xml";

import { createBlip } from "./blip";
import type { BlipOptions } from "./blip";
import type { BlipEffectsOptions } from "./blip-effects";
import { createSourceRectangle } from "./source-rectangle";
import type { SourceRectangleOptions } from "./source-rectangle";
import { createTileInfo } from "./tile";
import type { TileOptions } from "./tile";

/**
 * Options for blip fill properties.
 */
export interface BlipFillOptions {
  /** DPI of the image */
  dpi?: number;
  /** Whether the fill rotates with the shape */
  rotWithShape?: boolean;
  /** Image adjustment effects (brightness, contrast, grayscale, etc.) */
  blipEffects?: BlipEffectsOptions;
  /** Source rectangle for cropping */
  sourceRectangle?: SourceRectangleOptions;
  /** Tile fill mode (if omitted, defaults to stretch) */
  tile?: TileOptions;
}

/**
 * Creates a blip fill element.
 *
 * This element specifies the type of fill used for a picture. It contains the blip (image)
 * reference, an optional source rectangle for cropping, and the fill mode (typically stretch).
 *
 * Reference: ISO/IEC 29500-4, dml-main.xsd, CT_BlipFillProperties
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_BlipFillProperties">
 *   <xsd:sequence>
 *     <xsd:element name="blip" type="CT_Blip" minOccurs="0" maxOccurs="1"/>
 *     <xsd:element name="srcRect" type="CT_RelativeRect" minOccurs="0" maxOccurs="1"/>
 *     <xsd:group ref="EG_FillModeProperties" minOccurs="0" maxOccurs="1"/>
 *   </xsd:sequence>
 *   <xsd:attribute name="dpi" type="xsd:unsignedInt" use="optional"/>
 *   <xsd:attribute name="rotWithShape" type="xsd:boolean" use="optional"/>
 * </xsd:complexType>
 * ```
 *
 * @param blipOptions - Blip options including referenceId and type
 * @param fillOptions - Optional blip fill options
 * @returns An XML string representing the blip fill element
 */
export const createBlipFill = (blipOptions: BlipOptions, fillOptions?: BlipFillOptions): string => {
  const children: string[] = [];

  children.push(createBlip(blipOptions, fillOptions?.blipEffects));
  children.push(createSourceRectangle(fillOptions?.sourceRectangle));

  if (fillOptions?.tile) {
    children.push(createTileInfo(fillOptions.tile));
  } else {
    children.push("<a:stretch><a:fillRect/></a:stretch>");
  }

  const attrs: Record<string, string | number | undefined> = {};
  if (fillOptions?.dpi !== undefined) {
    attrs.dpi = fillOptions.dpi;
  }
  if (fillOptions?.rotWithShape !== undefined) {
    attrs.rotWithShape = fillOptions.rotWithShape ? 1 : 0;
  }

  return element("pic:blipFill", attrs, children);
};
