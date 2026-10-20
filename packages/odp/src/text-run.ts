import type {
  ParagraphDescriptorOptions,
  TextFieldOptions,
  TextBreakOptions,
  TextBodyOptions,
  TextRunOptions,
} from "@office-open/core";
import { attributeNumber, attributeString, escapeText, textOf, xmlElement } from "@office-open/odf";
import type { Element } from "@office-open/xml";

import { OdpGenerateError, unknownSlideChild } from "./error";
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
            if ("type" in run) return fieldXml(run as TextFieldOptions);
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

const OOXML_FIELD_TYPES: Record<string, TextFieldOptions["type"]> = {
  slidenum: "slidenum",
  slidecount: "slidecount",
  datetime: "datetime",
  datetimeFigureOut: "datetime",
  time: "time",
};

function odfFieldType(type: string): TextFieldOptions["type"] {
  if (OOXML_FIELD_TYPES[type]) return OOXML_FIELD_TYPES[type]!;
  if (type.startsWith("datetime")) return "datetime";
  if (type.startsWith("time")) return "time";
  return type as TextFieldOptions["type"];
}

function fieldXml(field: TextFieldOptions): string {
  const type = odfFieldType(field.type);
  const fixed = field.fixed === undefined ? undefined : String(field.fixed);
  switch (type) {
    case "slidenum":
      return xmlElement(
        "text:page-number",
        {
          "text:select-page": field.selectPage,
          "text:page-adjust": field.pageAdjust,
          "text:fixed": fixed,
        },
        [escapeText(field.text ?? "")],
      );
    case "slidecount":
      return xmlElement("text:page-count", { "text:fixed": fixed }, [escapeText(field.text ?? "")]);
    case "slidename":
      return xmlElement("loext:page-name", { "text:fixed": fixed }, [escapeText(field.text ?? "")]);
    case "author":
      return xmlElement("text:author-name", { "text:fixed": fixed }, [
        escapeText(field.text ?? ""),
      ]);
    case "filename":
      return xmlElement("text:file-name", { "text:display": field.display, "text:fixed": fixed }, [
        escapeText(field.text ?? ""),
      ]);
    case "datetime":
      return xmlElement(
        "text:date",
        {
          "style:data-style-name": field.dataStyleName,
          "text:date-value": field.dateValue,
          "text:fixed": fixed,
        },
        [escapeText(field.text ?? "")],
      );
    case "time":
      return xmlElement(
        "text:time",
        {
          "style:data-style-name": field.dataStyleName,
          "text:time-value": field.timeValue,
          "text:fixed": fixed,
        },
        [escapeText(field.text ?? "")],
      );
    default:
      throw new OdpGenerateError(
        `Unsupported text field type: ${field.type}`,
        "content.xml",
        "/text:p/field",
        field.type,
        "unsupported-text-field",
      );
  }
}

function parseField(element: Element): TextFieldOptions {
  const text = textOf(element);
  const fixed = attributeString(element, "text:fixed");
  const fixedOption = fixed === undefined ? {} : { fixed: fixed === "true" };
  switch (element.name) {
    case "text:page-number": {
      const pageAdjust = attributeNumber(element, "text:page-adjust");
      const selectPage = attributeString(element, "text:select-page");
      return {
        type: "slidenum",
        text,
        ...fixedOption,
        ...(selectPage === undefined
          ? {}
          : { selectPage: selectPage as "current" | "previous" | "next" }),
        ...(pageAdjust === undefined ? {} : { pageAdjust }),
      };
    }
    case "text:page-count":
      return { type: "slidecount", text, ...fixedOption };
    case "loext:page-name":
      return { type: "slidename", text, ...fixedOption };
    case "text:author-name":
      return { type: "author", text, ...fixedOption };
    case "text:file-name": {
      const display = attributeString(element, "text:display");
      return {
        type: "filename",
        text,
        ...fixedOption,
        ...(display === undefined
          ? {}
          : { display: display as "full" | "path" | "name" | "name-and-extension" }),
      };
    }
    case "text:date": {
      const dateValue = attributeString(element, "text:date-value");
      const dataStyleName = attributeString(element, "style:data-style-name");
      return {
        type: "datetime",
        text,
        ...fixedOption,
        ...(dateValue === undefined ? {} : { dateValue }),
        ...(dataStyleName === undefined ? {} : { dataStyleName }),
      };
    }
    case "text:time": {
      const timeValue = attributeString(element, "text:time-value");
      const dataStyleName = attributeString(element, "style:data-style-name");
      return {
        type: "time",
        text,
        ...fixedOption,
        ...(timeValue === undefined ? {} : { timeValue }),
        ...(dataStyleName === undefined ? {} : { dataStyleName }),
      };
    }
    default:
      throw unknownSlideChild(element, `/${element.name ?? "text:p"}`);
  }
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
    } else if (
      child.name === "text:page-number" ||
      child.name === "text:page-count" ||
      child.name === "loext:page-name" ||
      child.name === "text:author-name" ||
      child.name === "text:file-name" ||
      child.name === "text:date" ||
      child.name === "text:time"
    ) {
      children.push(parseField(child));
    } else if (child.name) {
      throw unknownSlideChild(child, `/${paragraph.name ?? "text:p"}`);
    }
  }
  if (children.length === 1 && typeof children[0] === "string") return { text: children[0] };
  return children.length ? { children } : { text: "" };
}
