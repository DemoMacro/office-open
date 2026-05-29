/**
 * XML traverser for searching paragraph content in OOXML documents.
 *
 * Parameterised by `XmlNamespaceConfig` so it works for both DOCX and PPTX.
 */
import type { Element } from "@office-open/xml";

import { createRunRenderer, type ElementWrapper, type RenderedParagraphNode } from "./run-renderer";
import type { XmlNamespaceConfig } from "./xml-namespace";

const elementsToWrapper = (wrapper: ElementWrapper): readonly ElementWrapper[] =>
  wrapper.element.elements?.map((e, i) => ({
    element: e,
    index: i,
    parent: wrapper,
  })) ?? [];

export function createTraverser(ns: XmlNamespaceConfig) {
  const renderParagraphNode = createRunRenderer(ns);

  const traverse = (node: Element): readonly RenderedParagraphNode[] => {
    let renderedParagraphs: readonly RenderedParagraphNode[] = [];

    const queue: ElementWrapper[] = [
      ...elementsToWrapper({
        element: node,
        index: 0,
        parent: undefined,
      }),
    ];

    let currentNode: ElementWrapper | undefined;
    while (queue.length > 0) {
      currentNode = queue.shift()!;

      if (currentNode.element.name === ns.paragraph) {
        renderedParagraphs = [...renderedParagraphs, renderParagraphNode(currentNode)];
      }
      queue.push(...elementsToWrapper(currentNode));
    }

    return renderedParagraphs;
  };

  const findLocationOfText = (node: Element, text: string): readonly RenderedParagraphNode[] =>
    traverse(node).filter((p) => p.text.includes(text));

  return { traverse, findLocationOfText };
}
