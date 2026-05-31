/**
 * XML traverser for searching paragraph content in OOXML documents.
 *
 * Parameterised by `XmlNamespaceConfig` so it works for both DOCX and PPTX.
 */
import type { Element } from "@office-open/xml";

import { createRunRenderer, type ElementWrapper, type RenderedParagraphNode } from "./run-renderer";
import type { XmlNamespaceConfig } from "./xml-namespace";

export function createTraverser(ns: XmlNamespaceConfig) {
  const renderParagraphNode = createRunRenderer(ns);

  const traverse = (node: Element): readonly RenderedParagraphNode[] => {
    const renderedParagraphs: RenderedParagraphNode[] = [];
    const queue: ElementWrapper[] = [];

    // Seed with root as virtual wrapper so paths include the root level
    const rootChildren = node.elements;
    if (rootChildren) {
      const root: ElementWrapper = { element: node, index: 0, parent: undefined };
      for (let i = 0; i < rootChildren.length; i++) {
        queue.push({ element: rootChildren[i], index: i, parent: root });
      }
    }

    let qi = 0;
    while (qi < queue.length) {
      const current = queue[qi++];

      if (current.element.name === ns.paragraph) {
        renderedParagraphs.push(renderParagraphNode(current));
      }

      const children = current.element.elements;
      if (children) {
        for (let i = 0; i < children.length; i++) {
          queue.push({ element: children[i], index: i, parent: current });
        }
      }
    }

    return renderedParagraphs;
  };

  const findLocationOfText = (node: Element, text: string): readonly RenderedParagraphNode[] =>
    traverse(node).filter((p) => p.text.includes(text));

  return { traverse, findLocationOfText };
}
