/**
 * Run renderer for extracting text content from OOXML paragraph structures.
 *
 * Parameterised by `XmlNamespaceConfig` so it works for both DOCX and PPTX.
 */
import type { Element } from "@office-open/xml";

import type { XmlNamespaceConfig } from "./xml-namespace";

export interface ElementWrapper {
  element: Element;
  index: number;
  parent: ElementWrapper | undefined;
}

export interface RenderedParagraphNode {
  text: string;
  runs: readonly RenderedRunNode[];
  index: number;
  pathToParagraph: readonly number[];
}

interface StartAndEnd {
  start: number;
  end: number;
}

type IParts = {
  text: string;
  index: number;
} & StartAndEnd;

export type RenderedRunNode = {
  text: string;
  parts: readonly IParts[];
  index: number;
} & StartAndEnd;

export function createRunRenderer(ns: XmlNamespaceConfig) {
  const renderParagraphNode = (node: ElementWrapper): RenderedParagraphNode => {
    if (node.element.name !== ns.paragraph) {
      throw new Error(`Invalid node type: ${node.element.name}`);
    }

    if (!node.element.elements) {
      return {
        index: -1,
        pathToParagraph: [],
        runs: [],
        text: "",
      };
    }

    let currentRunStringLength = 0;

    const runs: RenderedRunNode[] = [];
    for (let i = 0; i < node.element.elements.length; i++) {
      const element = node.element.elements[i];
      if (element.name === ns.run) {
        const renderedRunNode = renderRunNode(element, i, currentRunStringLength);
        currentRunStringLength += renderedRunNode.text.length;
        runs.push(renderedRunNode);
      }
    }

    const text = runs.reduce((acc, curr) => acc + curr.text, "");

    return {
      index: node.index,
      pathToParagraph: buildNodePath(node),
      runs,
      text,
    };
  };

  const renderRunNode = (
    node: Element,
    index: number,
    currentRunStringIndex: number,
  ): RenderedRunNode => {
    if (!node.elements) {
      return {
        end: currentRunStringIndex,
        index: -1,
        parts: [],
        start: currentRunStringIndex,
        text: "",
      };
    }

    let currentTextStringIndex = currentRunStringIndex;

    const parts: IParts[] = [];
    for (let i = 0; i < node.elements.length; i++) {
      const element = node.elements[i];
      if (element.name === ns.text && element.elements && element.elements.length > 0) {
        const text = element.elements[0].text ?? "";
        const textStr = String(text);
        const partStart = currentTextStringIndex;
        currentTextStringIndex += textStr.length;
        parts.push({ end: currentTextStringIndex - 1, index: i, start: partStart, text: textStr });
      }
    }

    const text = parts.reduce((acc, curr) => acc + curr.text, "");

    return {
      end: currentTextStringIndex - 1,
      index,
      parts,
      start: currentRunStringIndex,
      text,
    };
  };

  return renderParagraphNode;
}

const buildNodePath = (node: ElementWrapper): readonly number[] => {
  const path: number[] = [];
  let current: ElementWrapper | undefined = node;
  while (current) {
    path.push(current.index);
    current = current.parent;
  }
  path.reverse();
  return path;
};
