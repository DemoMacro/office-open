/**
 * External styles factory module for WordprocessingML documents.
 *
 * Parses styles from external XML and returns raw XML strings.
 * No XmlComponent dependency.
 *
 * Reference: http://officeopenxml.com/WPstyles.php
 *
 * @module
 */
import { js2xml, xml2js } from "@office-open/xml";
import type { Element as XMLElement } from "@office-open/xml";

import type { StylesOptions } from "./styles";

/**
 * Factory for creating styles from external XML sources.
 *
 * Parses styles from XML (typically from a styles.xml file)
 * and returns raw XML strings for each style element.
 */
export class ExternalStylesFactory {
  /**
   * Creates new Styles based on the given XML data.
   *
   * Parses the styles XML and converts each child to a raw XML string.
   */
  public newInstance(xmlData: string): StylesOptions {
    const xmlObj = xml2js(xmlData, { compact: false }) as XMLElement;

    let stylesXmlElement: XMLElement | undefined;
    for (const xmlElm of xmlObj.elements || []) {
      if (xmlElm.name === "w:styles") {
        stylesXmlElement = xmlElm;
      }
    }

    if (stylesXmlElement === undefined) {
      return { importedStyles: [], initialAttributes: {} };
    }

    const stylesElements = stylesXmlElement.elements || [];

    return {
      importedStyles: stylesElements.map((childElm) => ({
        _raw: js2xml({ elements: [childElm] }),
      })),
      initialAttributes: (stylesXmlElement.attributes as Record<string, string>) ?? {},
    };
  }
}
