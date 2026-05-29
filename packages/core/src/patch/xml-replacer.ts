/**
 * Parameterised XML replacer for placeholder substitution in OOXML documents.
 *
 * Accepts namespace config and a format callback so both DOCX and PPTX can
 * use the same replacement algorithm.
 */
import type { Element } from "@office-open/xml";

import { createSplitInject } from "./paragraph-split-inject";
import { createTokenReplacer } from "./paragraph-token-replacer";
import type { XmlNamespaceConfig } from "./xml-namespace";
import { createTextElementContents } from "./xml-patch-utils";
import { createTraverser } from "./xml-traverser";

export interface ReplacerConfig {
  readonly ns: XmlNamespaceConfig;
  readonly formatChild: (child: unknown, context: unknown) => Element[];
  readonly preserveSpace?: boolean;
}

interface ReplacerResult {
  readonly element: Element;
  readonly didFindOccurrence: boolean;
}

const SPLIT_TOKEN = "\u0275";

export function createReplacer(config: ReplacerConfig) {
  const { ns, formatChild } = config;
  const { findLocationOfText } = createTraverser(ns);

  const replaceTokenInParagraphElement = createTokenReplacer(createTextElementContents, {
    preserveSpace: config.preserveSpace,
  });
  const { findRunElementIndexWithToken, splitRunElement } = createSplitInject(
    ns,
    createTextElementContents,
    { preserveSpace: config.preserveSpace },
  );

  const replacer = ({
    json,
    patch,
    patchText,
    context,
    keepOriginalStyles = true,
  }: {
    readonly json: Element;
    readonly patch: {
      readonly type: string;
      readonly children: readonly unknown[];
    };
    readonly patchText: string;
    readonly context: unknown;
    readonly keepOriginalStyles?: boolean;
  }): ReplacerResult => {
    const renderedParagraphs = findLocationOfText(json, patchText);

    if (renderedParagraphs.length === 0) {
      return { didFindOccurrence: false, element: json };
    }

    for (const renderedParagraph of renderedParagraphs) {
      const textJson = patch.children.flatMap((c) => formatChild(c, context));

      switch (patch.type) {
        case "file": {
          const parentElement = goToParentElementFromPath(json, renderedParagraph.pathToParagraph);
          const elementIndex = getLastElementIndexFromPath(renderedParagraph.pathToParagraph);
          parentElement.elements!.splice(elementIndex, 1, ...textJson);
          break;
        }
        case "paragraph":
        default: {
          const paragraphElement = goToElementFromPath(json, renderedParagraph.pathToParagraph);
          replaceTokenInParagraphElement({
            originalText: patchText,
            paragraphElement,
            renderedParagraph,
            replacementText: SPLIT_TOKEN,
          });

          const index = findRunElementIndexWithToken(paragraphElement, SPLIT_TOKEN);

          const runElementToBeReplaced = paragraphElement.elements![index];
          const { left, right } = splitRunElement(runElementToBeReplaced, SPLIT_TOKEN);

          let newRunElements = textJson;
          let patchedRightElement = right;

          if (keepOriginalStyles) {
            const runElementNonTextualElements = runElementToBeReplaced.elements!.filter(
              (e) => e.type === "element" && e.name === ns.runProperties,
            );

            newRunElements = textJson.map((e) => {
              if (
                e.type !== "element" ||
                e.name !== ns.run ||
                e.elements?.some((c: any) => c.type === "element" && c.name === ns.runProperties)
              ) {
                return e;
              }
              return {
                ...e,
                elements: [...runElementNonTextualElements, ...(e.elements ?? [])],
              };
            });

            patchedRightElement = {
              ...right,
              elements: [...runElementNonTextualElements, ...right.elements!],
            };
          }

          paragraphElement.elements!.splice(index, 1, left, ...newRunElements, patchedRightElement);
          break;
        }
      }
    }

    return { didFindOccurrence: true, element: json };
  };

  return replacer;
}

const goToElementFromPath = (json: Element, path: readonly number[]): Element => {
  let element = json;
  for (let i = 1; i < path.length; i++) {
    const index = path[i];
    const nextElements = element.elements!;
    element = nextElements[index];
  }
  return element;
};

const goToParentElementFromPath = (json: Element, path: readonly number[]): Element =>
  goToElementFromPath(json, path.slice(0, -1));

const getLastElementIndexFromPath = (path: readonly number[]): number => path[path.length - 1];
