/**
 * Paragraph split and inject utilities for run-level text replacement.
 *
 * Parameterised by `XmlNamespaceConfig` so it works for both DOCX and PPTX.
 */
import type { Element } from "@office-open/xml";

import type { XmlNamespaceConfig } from "./xml-namespace";
import { patchSpaceAttribute } from "./xml-patch-utils";

export class TokenNotFoundError extends Error {
  public constructor(token: string) {
    super(`Token ${token} not found`);
    this.name = "TokenNotFoundError";
  }
}

export function createSplitInject(
  ns: XmlNamespaceConfig,
  createTextElementContents: (text: string) => Element[],
  options?: { readonly preserveSpace?: boolean },
) {
  const preserveSpace = options?.preserveSpace ?? true;
  const findRunElementIndexWithToken = (paragraphElement: Element, token: string): number => {
    for (let i = 0; i < (paragraphElement.elements ?? []).length; i++) {
      const element = paragraphElement.elements![i];
      if (element.type === "element" && element.name === ns.run) {
        const textElement = (element.elements ?? []).filter(
          (e) => e.type === "element" && e.name === ns.text,
        );

        for (const text of textElement) {
          if (!text.elements?.[0]) {
            continue;
          }

          if ((text.elements[0].text as string)?.includes(token)) {
            return i;
          }
        }
      }
    }

    throw new TokenNotFoundError(token);
  };

  const splitRunElement = (
    runElement: Element,
    token: string,
  ): { readonly left: Element; readonly right: Element } => {
    let splitIndex = -1;

    const splitElements =
      runElement.elements
        ?.map((e, i) => {
          if (splitIndex !== -1) {
            return e;
          }

          if (e.type === "element" && e.name === ns.text) {
            const text = (e.elements?.[0]?.text as string) ?? "";
            const splitText = text.split(token);
            const newElements = splitText.map((t) => ({
              ...e,
              ...(preserveSpace ? patchSpaceAttribute(e) : {}),
              elements: createTextElementContents(t),
            }));

            if (splitText.length > 1) {
              splitIndex = i;
            }

            return newElements;
          } else {
            return e;
          }
        })
        .flat() ?? [];

    const leftRunElement: Element = {
      ...runElement,
      elements: splitElements.slice(0, splitIndex + 1),
    };

    const rightRunElement: Element = {
      ...runElement,
      elements: splitElements.slice(splitIndex + 1),
    };

    return { left: leftRunElement, right: rightRunElement };
  };

  return { findRunElementIndexWithToken, splitRunElement };
}
