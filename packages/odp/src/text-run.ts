import type {
  ParagraphDescriptorOptions,
  TextBreakOptions,
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
            if (typeof run === "string") return escapeText(run);
            if ("break" in run) return xmlElement("text:line-break");
            const textRun = run as TextRunOptions;
            if (!("text" in run) || !textRun.hyperlink) return "";
            const properties = textProperties(textRun);
            const styleName = addTextStyle(properties, styles);
            const span = xmlElement("text:span", { "text:style-name": styleName }, [
              escapeText(textRun.text ?? ""),
            ]);
            return textRun.hyperlink.url
              ? xmlElement("text:a", { "xlink:href": textRun.hyperlink.url }, [span])
              : span;
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
  const children: (string | TextRunOptions | TextBreakOptions)[] = [];
  for (const child of paragraph.elements ?? []) {
    if (child.type === "text") children.push(String(child.text ?? ""));
    else if (child.name === "text:line-break")
      children.push({ break: true } satisfies TextBreakOptions);
    else if (child.name === "text:a") {
      const url = attributeString(child, "xlink:href");
      for (const run of child.elements ?? []) {
        if (run.type === "text")
          children.push({ text: String(run.text ?? ""), hyperlink: { url } });
        else if (run.name === "text:span")
          children.push({
            text: textOf(run),
            ...textStyles.get(attributeString(run, "text:style-name") ?? ""),
            hyperlink: { url },
          });
        else throw unknownSlideChild(child, `/${paragraph.name ?? "text:p"}/text:a`);
      }
    } else if (child.name === "text:s") {
      const count = Number(attributeString(child, "text:c") ?? "1");
      children.push(" ".repeat(Number.isInteger(count) && count > 0 ? count : 1));
    } else if (child.name === "text:tab") children.push({ text: "\t" });
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
