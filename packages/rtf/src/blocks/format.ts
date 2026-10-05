import type { RunOptions } from "@office-open/docx";

import type { RunFormat } from "./models";

export const CONTENT_GROUP_CONTROLS = new Set([
  "b",
  "cell",
  "cellx",
  "cf",
  "f",
  "fs",
  "i",
  "line",
  "nestcell",
  "nestrow",
  "nosupersub",
  "page",
  "par",
  "pard",
  "plain",
  "qc",
  "qj",
  "ql",
  "qr",
  "s",
  "strike",
  "sub",
  "super",
  "row",
  "sect",
  "tab",
  "trowd",
  "u",
  "uc",
  "ul",
  "ulnone",
  "ulw",
]);

export function isFormatChanged(left: RunFormat, right: RunFormat): boolean {
  return (
    ["bold", "italic", "underline", "color", "size", "font", "strike", "verticalAlign"] as const
  ).some((key) => left[key] !== right[key]);
}

export function runOptions(format: RunFormat): RunOptions {
  return { ...format };
}
