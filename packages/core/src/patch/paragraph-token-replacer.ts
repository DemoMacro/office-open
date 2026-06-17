/**
 * Paragraph token replacer for substituting text within paragraph runs.
 *
 * Fully generic — operates on Element nodes without namespace-specific logic.
 */
import type { Element } from "@office-open/xml";

import type { RenderedParagraphNode } from "./run-renderer";
import { patchSpaceAttribute } from "./xml-patch-utils";

const ReplaceMode = {
  START: 0,
  MIDDLE: 1,
  END: 2,
} as const;

export function createTokenReplacer(
  createTextElementContents: (text: string) => Element[],
  options?: { readonly preserveSpace?: boolean },
) {
  const preserveSpace = options?.preserveSpace ?? true;

  const patchTextElement = (element: Element, text: string): Element => {
    element.elements = createTextElementContents(text);
    return element;
  };

  return ({
    paragraphElement,
    renderedParagraph,
    originalText,
    replacementText,
  }: {
    paragraphElement: Element;
    renderedParagraph: RenderedParagraphNode;
    originalText: string;
    replacementText: string;
  }): Element => {
    const startIndex = renderedParagraph.text.indexOf(originalText);
    const endIndex = startIndex + originalText.length - 1;

    let replaceMode: (typeof ReplaceMode)[keyof typeof ReplaceMode] = ReplaceMode.START;

    for (const run of renderedParagraph.runs) {
      for (const { text, index, start, end } of run.parts) {
        switch (replaceMode) {
          case ReplaceMode.START: {
            if (startIndex >= start && startIndex <= end) {
              const offsetStartIndex = startIndex - start;
              const offsetEndIndex = Math.min(endIndex, end) - start;
              const partToReplace = text.substring(offsetStartIndex, offsetEndIndex + 1);
              if (partToReplace === "") {
                continue;
              }

              const firstPart = text.replace(partToReplace, replacementText);
              patchTextElement(paragraphElement.elements![run.index].elements![index], firstPart);
              replaceMode = ReplaceMode.MIDDLE;
              continue;
            }
            break;
          }
          case ReplaceMode.MIDDLE: {
            if (endIndex <= end) {
              const lastPart = text.substring(endIndex - start + 1);
              patchTextElement(paragraphElement.elements![run.index].elements![index], lastPart);
              const currentElement = paragraphElement.elements![run.index].elements![index];
              paragraphElement.elements![run.index].elements![index] = preserveSpace
                ? patchSpaceAttribute(currentElement)
                : currentElement;
              replaceMode = ReplaceMode.END;
            } else {
              patchTextElement(paragraphElement.elements![run.index].elements![index], "");
            }
            break;
          }
        }
      }
    }

    return paragraphElement;
  };
}
