import { escapeText, xmlElement } from "@office-open/odf";
import type { CommentOptions, WorksheetOptions } from "@office-open/xlsx";

export function commentPlainText(text: CommentOptions["text"]): string {
  if (typeof text === "string") return text;
  return text.runs?.map((run) => run.text).join("") ?? text.text ?? "";
}

export function cellDecorationXml(worksheet: WorksheetOptions, reference: string): string[] {
  const comment = worksheet.comments?.find((entry) => entry.cell === reference);
  return comment
    ? [
        xmlElement(
          "office:annotation",
          {
            "dc:creator": comment.author,
          },
          [xmlElement("text:p", undefined, [escapeText(commentPlainText(comment.text))])],
        ),
      ]
    : [];
}
