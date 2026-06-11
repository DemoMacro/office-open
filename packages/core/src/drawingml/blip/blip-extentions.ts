/**
 * Blip extensions module for SVG support.
 *
 * This module provides extension elements that enable SVG image support
 * within blip elements using Office-specific extensions.
 *
 * Reference: http://officeopenxml.com/drwPic.php
 *
 * @module
 */
import { element } from "@office-open/xml";

/**
 * Creates an SVG blip element for embedding SVG images.
 *
 * This element is a Microsoft Office extension that allows SVG images
 * to be referenced within a blip.
 *
 * @param svgReferenceId - The reference ID for the SVG image
 * @returns An XML string representing the SVG blip element
 * @internal
 */
const createSvgBlip = (svgReferenceId: string): string =>
  `<asvg:svgBlip xmlns:asvg="http://schemas.microsoft.com/office/drawing/2016/SVG/main" r:embed="{${svgReferenceId}}"/>`;

/**
 * Creates an extension element for SVG support.
 *
 * This element wraps the SVG blip extension with the appropriate URI
 * to identify it as an SVG extension.
 *
 * @param svgReferenceId - The reference ID for the SVG image
 * @returns An XML string representing the extension element
 * @internal
 */
const createExtention = (svgReferenceId: string): string =>
  element("a:ext", { uri: "{96DAC541-7B7A-43D3-8B79-37D633B846F1}" }, [
    createSvgBlip(svgReferenceId),
  ]);

/**
 * Creates an extension list for SVG images.
 *
 * This element contains the extensions needed to embed SVG images
 * within a blip. It wraps the SVG-specific extension elements.
 *
 * ## XSD Schema
 * ```xml
 * <xsd:complexType name="CT_OfficeArtExtensionList">
 *   <xsd:sequence>
 *     <xsd:element name="ext" type="CT_OfficeArtExtension" minOccurs="0" maxOccurs="unbounded"/>
 *   </xsd:sequence>
 * </xsd:complexType>
 * ```
 *
 * @param svgReferenceId - The reference ID for the SVG image
 * @returns An XML string representing the extension list
 */
export const createExtentionList = (svgReferenceId: string): string =>
  element("a:extLst", undefined, [createExtention(svgReferenceId)]);
