import type { ParagraphOptions } from "@office-open/docx";

import type { RunFormat } from "../blocks";

export type Destination = "root" | "font-table" | "color-table" | "generated-shadow";

export type GroupFrame = {
  destination: Destination;
  startIndex: number;
  starred: boolean;
  decided: boolean;
  skip: boolean;
  format: RunFormat;
  alignment: ParagraphOptions["alignment"];
  paragraphAlignment: ParagraphOptions["alignment"];
  fontId?: number;
  fontName?: string;
};
