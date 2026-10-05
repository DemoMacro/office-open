import type { TableOptions } from "@office-open/pptx";

import type { RecordNode } from "../records";
export interface TextEntry {
  readonly placeholderType: number;
  readonly text: string;
}

export type TableCell = NonNullable<TableOptions["rows"]>[number]["cells"][number];

export interface SlideTextGroup {
  readonly persistReference: number;
  readonly slideIdentifier: number | undefined;
  readonly entries: TextEntry[];
}

export interface LegacyPicture {
  readonly type: "png" | "jpg" | "gif" | "bmp" | "emf" | "wmf";
  readonly data: Uint8Array;
}

export interface DrawingContext {
  readonly entries: TextEntry[];
  readonly pictures: readonly (LegacyPicture | undefined)[];
  readonly usedEntryIndexes: Set<number>;
  readonly hyperlinks: ReadonlyMap<number, { url?: string; tooltip?: string }>;
}

export interface LegacyTextHyperlink {
  readonly start: number;
  readonly end: number;
  readonly target: { url?: string; tooltip?: string };
}

export interface LegacyAnimation {
  readonly shapeId: number | undefined;
  readonly atom: RecordNode;
}
