import type {
  ParagraphDescriptorOptions,
  TextBodyOptions,
  TextRunOptions,
} from "@office-open/core";
import { attributeString, escapeText, textOf, xmlElement } from "@office-open/odf";
import type { Element } from "@office-open/xml";

import { unknownSlideChild } from "./presentation";
import { addTextStyle } from "./styles";

export interface TextProperties {
  bold?: boolean;
  italic?: boolean;
  underline?: "single";
  size?: number;
}

export function textBodyXml(body: TextBodyOptions | undefined, styles: string[]): string[] {
  const paragraphs = body?.paragraphs ?? (body?.text !== undefined ? [body.text] : [""]);
  return paragraphs.map((paragraph) => {
    const value = typeof paragraph === "string" ? { text: paragraph } : paragraph;
    const children =
      value.text !== undefined && value.children === undefined
        ? [escapeText(value.text)]
        : (value.children ?? []).map((run) => {
            if (typeof run === "string")
              return xmlElement("text:span", undefined, [escapeText(run)]);
            if (!("text" in run)) return "";
            const properties = textProperties(run);
            const styleName = addTextStyle(properties, styles);
            return xmlElement("text:span", { "text:style-name": styleName }, [
              escapeText(run.text ?? ""),
            ]);
          });
    return xmlElement("text:p", undefined, children);
  });
}

export function textProperties(run: TextRunOptions): TextProperties {
  return {
    bold: run.bold,
    italic: run.italic,
    underline: run.underline === "single" ? "single" : undefined,
    size: run.size,
  };
}

export function parseParagraph(
  paragraph: Element,
  textStyles: Map<string, TextProperties>,
): ParagraphDescriptorOptions {
  const children: (string | TextRunOptions)[] = [];
  for (const child of paragraph.elements ?? []) {
    if (child.type === "text") children.push(String(child.text ?? ""));
    else if (child.name === "text:span") {
      children.push({
        text: textOf(child),
        ...textStyles.get(attributeString(child, "text:style-name") ?? ""),
      });
    } else if (child.name) {
      throw unknownSlideChild(child, `/${paragraph.name ?? "text:p"}`);
    }
  }
  if (children.length === 1 && typeof children[0] === "string") return { text: children[0] };
  return children.length ? { children } : { text: "" };
}
