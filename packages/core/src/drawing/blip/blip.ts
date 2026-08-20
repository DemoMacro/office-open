/**
 * Blip (Binary Large Image or Picture) module for DrawingML.
 *
 * This module provides the blip element that references the actual
 * image data within a picture.
 *
 * Reference: http://officeopenxml.com/drwPic.php
 *
 * @module
 */
import { element } from "@office-open/xml";

import { createBlipEffects } from "./blip-effects";
import type { BlipEffectsOptions } from "./blip-effects";
import { createExtensionList } from "./blip-extensions";

/** Blip compression states (ST_BlipCompression, the a:blip @cstate attribute). */
export type BlipCompression = "email" | "screen" | "print" | "hqprint" | "none";

/**
 * Options for creating a blip element.
 */
export interface BlipOptions {
  /**
   * File name used as placeholder for the embedded image; the packer's
   * ImageReplacer replaces `{referenceId}` with `rId{N}`. Absent on a
   * linked-only picture (external URL, no bytes in the package).
   */
  referenceId?: string;
  /**
   * Placeholder key for the linked image source (a:blip @r:link) — an external
   * URL relationship. Present alone on linked-only pictures, alongside
   * referenceId when the package also caches a local copy.
   */
  linkReferenceId?: string;
  /** Compression state (a:blip @cstate); absent = attribute omitted (schema default "none"). */
  compression?: BlipCompression;
  /** Image type for SVG detection */
  type?: "svg" | string;
  /** For SVG images, the fallback image file name */
  fallbackFileName?: string;
}

/**
 * Creates a blip element for an image.
 *
 * A blip references the actual image data stored in the document package
 * through a relationship ID. For SVG images, it includes extensions that
 * reference the SVG data.
 *
 * Reference: http://officeopenxml.com/drwPic.php
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_Blip">
 *   <xsd:sequence>
 *     <xsd:choice minOccurs="0" maxOccurs="unbounded">
 *       <xsd:element name="alphaBiLevel" type="CT_AlphaBiLevelEffect"/>
 *       <xsd:element name="alphaCeiling" type="CT_AlphaCeilingEffect"/>
 *       <xsd:element name="alphaFloor" type="CT_AlphaFloorEffect"/>
 *       <xsd:element name="alphaInv" type="CT_AlphaInverseEffect"/>
 *       <xsd:element name="alphaMod" type="CT_AlphaModulateEffect"/>
 *       <xsd:element name="alphaModFix" type="CT_AlphaModulateFixedEffect"/>
 *       <xsd:element name="alphaRepl" type="CT_AlphaReplaceEffect"/>
 *       <xsd:element name="biLevel" type="CT_BiLevelEffect"/>
 *       <xsd:element name="blur" type="CT_BlurEffect"/>
 *       <xsd:element name="clrChange" type="CT_ColorChangeEffect"/>
 *       <xsd:element name="clrRepl" type="CT_ColorReplaceEffect"/>
 *       <xsd:element name="duotone" type="CT_DuotoneEffect"/>
 *       <xsd:element name="fillOverlay" type="CT_FillOverlayEffect"/>
 *       <xsd:element name="grayscl" type="CT_GrayscaleEffect"/>
 *       <xsd:element name="hsl" type="CT_HSLEffect"/>
 *       <xsd:element name="lum" type="CT_LuminanceEffect"/>
 *       <xsd:element name="tint" type="CT_TintEffect"/>
 *     </xsd:choice>
 *     <xsd:element name="extLst" type="CT_OfficeArtExtensionList" minOccurs="0"/>
 *   </xsd:sequence>
 *   <xsd:attribute ref="r:embed"/>
 *   <xsd:attribute ref="r:link"/>
 *   <xsd:attribute name="cstate" type="ST_BlipCompression"/>
 * </xsd:complexType>
 * ```
 *
 * @param options - Blip options including referenceId and type
 * @param blipEffects - Optional blip effects (brightness, contrast, etc.)
 * @returns An XML string representing the blip element
 */
export const createBlip = (options: BlipOptions, blipEffects?: BlipEffectsOptions): string => {
  const children: string[] = [];

  if (blipEffects) {
    children.push(...createBlipEffects(blipEffects));
  }

  if (options.type === "svg" && options.fallbackFileName && options.referenceId !== undefined) {
    children.push(createExtensionList(options.referenceId));
  }

  return element(
    "a:blip",
    {
      ...(options.referenceId !== undefined
        ? {
            "r:embed": `{${options.type === "svg" && options.fallbackFileName ? options.fallbackFileName : options.referenceId}}`,
          }
        : {}),
      ...(options.compression !== undefined ? { cstate: options.compression } : {}),
      ...(options.linkReferenceId !== undefined
        ? { "r:link": `{${options.linkReferenceId}}` }
        : {}),
    },
    children.length > 0 ? children : undefined,
  );
};
