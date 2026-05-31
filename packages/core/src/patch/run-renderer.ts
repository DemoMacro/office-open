/**
 * Run renderer for extracting text content from OOXML paragraph structures.
 *
 * Parameterised by `XmlNamespaceConfig` so it works for both DOCX and PPTX.
 */
import type { Element } from "@office-open/xml";

import type { XmlNamespaceConfig } from "./xml-namespace";

export interface ElementWrapper {
  readonly element: Element;
  readonly index: number;
  readonly parent: ElementWrapper | undefined;
}

export interface RenderedParagraphNode {
  readonly text: string;
  readonly runs: readonly IRenderedRunNode[];
  readonly index: number;
  readonly pathToParagraph: readonly number[];
}

interface StartAndEnd {
  readonly start: number;
  readonly end: number;
}

type IParts = {
  readonly text: string;
  readonly index: number;
} & StartAndEnd;

export type IRenderedRunNode = {
  readonly text: string;
  readonly parts: readonly IParts[];
  readonly index: number;
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

    const runs = node.element.elements
      .map((element, i) => ({ element, i }))
      .filter(({ element }) => element.name === ns.run)
      .map(({ element, i }) => {
        const renderedRunNode = renderRunNode(element, i, currentRunStringLength);
        currentRunStringLength += renderedRunNode.text.length;
        return renderedRunNode;
      })
      .filter((e) => Boolean(e));

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
  ): IRenderedRunNode => {
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

    const parts = node.elements
      .map((element, i: number) =>
        element.name === ns.text && element.elements && element.elements.length > 0
          ? (() => {
              const partStart = currentTextStringIndex;
              currentTextStringIndex += (element.elements[0].text?.toString() ?? "").length;
              return {
                end: currentTextStringIndex - 1,
                index: i,
                start: partStart,
                text: element.elements[0].text?.toString() ?? "",
              };
            })()
          : undefined,
      )
      .filter((e) => Boolean(e))
      .map((e) => e as IParts);

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
