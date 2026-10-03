import type {
  DocumentOptions,
  ParagraphOptions,
  RunOptions,
  SectionChild,
  TableCellOptions,
  TableRowOptions,
  TableOptions,
} from "@office-open/docx";

import { RtfParseError } from "./errors";
import { tokenizeRtf, type RtfToken } from "./tokenizer";

type RunFormat = Pick<RunOptions, "bold" | "italic" | "underline" | "color" | "size" | "font">;
type ParagraphDraft = {
  options: ParagraphOptions;
  runs: RunOptions[];
  format: RunFormat;
  segments: (string | NonNullable<RunOptions["children"]>[number])[];
};
type CellDraft = { blocks: SectionChild[] };
type TableDraft = {
  target: TableOptions;
  rows: TableRowOptions[];
  cells: TableCellOptions[];
  cell: CellDraft;
  cellOpen: boolean;
  cellEnded: boolean;
  parentBlocks: SectionChild[];
};
type Destination = "root" | "font-table" | "color-table" | "ignored";
type GroupFrame = {
  destination: Destination;
  starred: boolean;
  decided: boolean;
  skip: boolean;
  format: RunFormat;
  alignment: ParagraphOptions["alignment"];
  paragraphAlignment: ParagraphOptions["alignment"];
  fontId?: number;
  fontName?: string;
};

const IGNORED_DESTINATIONS = new Set([
  "author",
  "bkmkend",
  "bkmkstart",
  "category",
  "company",
  "datafield",
  "datastore",
  "doccomm",
  "falttext",
  "field",
  "filetbl",
  "footer",
  "footerf",
  "footerl",
  "footerr",
  "footnote",
  "header",
  "headerf",
  "headerl",
  "headerr",
  "hlinkbase",
  "hlink",
  "info",
  "keywords",
  "listtable",
  "listoverridetable",
  "manager",
  "objdata",
  "operator",
  "pict",
  "pn",
  "pntext",
  "proto",
  "revtbl",
  "rsidgrp",
  "rsidtbl",
  "shpinst",
  "stylesheet",
  "subject",
  "themedata",
  "title",
  "xmlnstb",
]);

const CONTENT_GROUP_CONTROLS = new Set([
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
  "page",
  "par",
  "pard",
  "plain",
  "qc",
  "qj",
  "ql",
  "qr",
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

function isFormatChanged(left: RunFormat, right: RunFormat): boolean {
  return (["bold", "italic", "underline", "color", "size", "font"] as const).some(
    (key) => left[key] !== right[key],
  );
}

function runOptions(format: RunFormat): RunOptions {
  return { ...format };
}

export function parseRtf(source: string): DocumentOptions {
  const tokens = tokenizeRtf(source);
  if (
    tokens[0]?.kind !== "group-start" ||
    tokens[1]?.kind !== "control" ||
    tokens[1].word !== "rtf"
  ) {
    throw new RtfParseError('RTF must begin with "{\\rtf"', tokens[0]?.position ?? 0, source);
  }
  const blocks: SectionChild[] = [];
  const fonts = new Map<number, string>();
  const colors: (string | null)[] = [];
  const tables: TableDraft[] = [];
  const groupFrames: GroupFrame[] = [];

  let format: RunFormat = {};
  let alignment: ParagraphOptions["alignment"];
  let paragraph: ParagraphDraft | undefined;
  let unicodeByteCount = 1;
  let unicodeSkip = 0;
  let fontId: number | undefined;
  let fontName: string | undefined;
  let color: { red?: number; green?: number; blue?: number } = {};
  let tokenIndex = 0;

  const activeTable = () => tables.at(-1);

  const closeRow = (draft: TableDraft) => {
    if (!draft.cellEnded) finalizeCell(draft);
    draft.target.rows.push({ cells: draft.cells });
    draft.cells = [];
    draft.cellEnded = false;
    tables.pop();
    paragraph = undefined;
  };

  const ensureParagraph = () => {
    if (paragraph) return paragraph;
    paragraph = { options: {}, runs: [], format: { ...format }, segments: [] };
    return paragraph;
  };

  const flushRun = (draft: ParagraphDraft | undefined) => {
    if (!draft) return;
    const hasSegments = draft.segments.length > 0;
    if (!hasSegments) return;
    const run = runOptions(draft.format);
    if (draft.segments.length === 1 && typeof draft.segments[0] === "string") {
      run.text = draft.segments[0];
    } else {
      run.children = draft.segments.map((segment) =>
        typeof segment === "string" ? { text: segment } : segment,
      );
    }
    draft.runs.push(run);
    draft.segments = [];
  };

  const flushParagraph = (required = false) => {
    const draft = paragraph;
    if (!draft) {
      if (required) blocks.push({ paragraph: {} });
      return;
    }
    flushRun(draft);
    const hasAlignment = draft.options.alignment !== undefined;
    if (draft.runs.length > 0 || hasAlignment || required) {
      const options = { ...draft.options };
      if (draft.runs.length > 0) options.children = draft.runs;
      (activeTable()?.cell.blocks ?? blocks).push({ paragraph: options });
    }
    paragraph = undefined;
  };

  const startParagraph = () => {
    flushParagraph();
    ensureParagraph();
  };

  const appendText = (value: string) => {
    const draft = ensureParagraph();
    if (draft.format !== format && isFormatChanged(draft.format, format)) {
      flushRun(draft);
      draft.format = { ...format };
    }
    const previous = draft.segments.at(-1);
    if (typeof previous === "string") draft.segments[draft.segments.length - 1] = previous + value;
    else draft.segments.push(value);
  };

  const appendInline = (child: NonNullable<RunOptions["children"]>[number]) => {
    const draft = ensureParagraph();
    flushRun(draft);
    draft.format = { ...format };
    draft.segments.push(child);
  };

  const ensureTable = () => {
    const active = activeTable();
    if (active?.cellEnded) closeRow(active);
    const current = activeTable();
    const parentBlocks = current ? current.cell.blocks : blocks;
    flushParagraph();
    const previous = parentBlocks.at(-1);
    const target = previous && "table" in previous ? previous.table : { rows: [] };
    if (!previous || !("table" in previous)) parentBlocks.push({ table: target });
    tables.push({
      target,
      rows: [],
      cells: [],
      cell: { blocks: [] },
      cellOpen: true,
      cellEnded: false,
      parentBlocks,
    });
    paragraph = undefined;
  };

  const finalizeCell = (draft: TableDraft) => {
    flushParagraph();
    if (draft.cellOpen) {
      draft.cells.push({ children: draft.cell.blocks });
      draft.cellOpen = false;
      draft.cell = { blocks: [] };
    }
  };

  const endRow = () => {
    const draft = activeTable();
    if (!draft) return;
    closeRow(draft);
  };

  const addCell = () => {
    const draft = activeTable();
    if (!draft) {
      ensureTable();
      return;
    }
    finalizeCell(draft);
    draft.cellOpen = true;
    draft.cellEnded = true;
  };

  const endNestedRow = () => {
    const draft = activeTable();
    if (draft) closeRow(draft);
  };

  const decodeUnicode = (value: number) => {
    const codePoint = value < 0 ? value + 65536 : value;
    if (!Number.isInteger(codePoint) || codePoint < 0 || codePoint > 0x10ffff) {
      throw new RtfParseError(
        `\\u value ${value} is not a valid character code`,
        tokens[tokenIndex]?.position ?? 0,
        source,
      );
    }
    appendText(String.fromCodePoint(codePoint));
  };

  const parseFontTableControl = (word: string, parameter?: number) => {
    if (word === "f") {
      fontId = parameter;
      fontName = "";
      return true;
    }
    if (
      word === "fcharset" ||
      word === "froman" ||
      word === "fmodern" ||
      word === "fswiss" ||
      word === "fdecor" ||
      word === "fscript" ||
      word === "ftech" ||
      word === "fbidi"
    ) {
      return true;
    }
    return false;
  };

  const parseColorControl = (word: string, parameter?: number) => {
    if (word === "red") color.red = parameter;
    else if (word === "green") color.green = parameter;
    else if (word === "blue") color.blue = parameter;
    else return false;
    return true;
  };

  const pushColor = () => {
    if (color.red === undefined && color.green === undefined && color.blue === undefined) {
      colors.push(null);
    } else {
      const component = (value: number | undefined) =>
        Math.max(0, Math.min(255, value ?? 0))
          .toString(16)
          .padStart(2, "0")
          .toUpperCase();
      colors.push(`${component(color.red)}${component(color.green)}${component(color.blue)}`);
    }
    color = {};
  };

  const applyControl = (
    token: Extract<RtfToken, { kind: "control" }>,
    destination: Destination,
  ) => {
    const word = token.word;
    if (!word) {
      if (token.symbol === "*") return;
      if (token.symbol === "~") return appendText(" ");
      if (token.symbol === "-") return appendText("");
      if (token.symbol === "_") return appendText("‐");
      if (token.symbol === "{") return appendText("{");
      if (token.symbol === "}") return appendText("}");
      if (token.symbol === "\\") return appendText("\\");
      return;
    }

    if (destination === "font-table" && parseFontTableControl(word, token.param)) return;
    if (destination === "color-table" && parseColorControl(word, token.param)) return;

    switch (word) {
      case "par":
        startParagraph();
        return;
      case "tab":
        appendInline({ tab: true });
        return;
      case "line":
      case "page":
      case "sect":
        appendInline({ break: 1 });
        return;
      case "u":
        if (token.param === undefined)
          throw new RtfParseError("\\u requires a character code", token.position, source);
        decodeUnicode(token.param);
        unicodeSkip = unicodeByteCount;
        return;
      case "uc":
        unicodeByteCount = token.param ?? 1;
        return;
      case "plain":
        format = {};
        return;
      case "pard":
        setAlignment(undefined);
        return;
      case "ql":
        setAlignment("left");
        return;
      case "qr":
        setAlignment("right");
        return;
      case "qc":
        setAlignment("center");
        return;
      case "qj":
        setAlignment("both");
        return;
      case "b":
        format.bold = token.param !== 0;
        return;
      case "i":
        format.italic = token.param !== 0;
        return;
      case "ul":
        format.underline = { type: "single" };
        return;
      case "ulnone":
        format.underline = undefined;
        return;
      case "ulw":
        format.underline = { type: "single" };
        return;
      case "fs":
        if (!token.param || token.param <= 0)
          throw new RtfParseError(
            "\\fs requires a positive half-point size",
            token.position,
            source,
          );
        format.size = token.param / 2;
        return;
      case "f":
        format.font = token.param !== undefined ? fonts.get(token.param) : undefined;
        return;
      case "cf": {
        const mapped = token.param !== undefined ? colors[token.param] : undefined;
        format.color = mapped ?? undefined;
        return;
      }
      case "trowd":
        ensureTable();
        return;
      case "cell":
      case "nestcell":
        addCell();
        return;
      case "row":
        endRow();
        return;
      case "nestrow":
        endNestedRow();
        return;
      case "cellx":
        if (token.param === undefined)
          throw new RtfParseError("\\cellx requires a boundary", token.position, source);
        const table = activeTable()?.target;
        if (!table) return;
        table.columnWidths ??= [];
        table.columnWidths.push(token.param);
        return;
      default:
        return;
    }
  };

  const setAlignment = (value: ParagraphOptions["alignment"]) => {
    alignment = value;
    if (paragraph) paragraph.options.alignment = value;
  };

  const handleGroupStart = (position: number) => {
    groupFrames.push({
      destination: groupFrames.at(-1)?.destination ?? "root",
      starred: false,
      decided: false,
      skip: false,
      format: { ...format },
      alignment,
      paragraphAlignment: paragraph?.options.alignment,
      fontId,
      fontName,
    });
    void position;
  };

  const handleGroupEnd = () => {
    const frame = groupFrames.pop();
    if (!frame) return;
    if (
      frame.destination === "color-table" &&
      (color.red !== undefined || color.green !== undefined || color.blue !== undefined)
    )
      pushColor();
    format = frame.format;
    alignment = frame.alignment;
    if (paragraph) {
      if (frame.paragraphAlignment === undefined) delete paragraph.options.alignment;
      else paragraph.options.alignment = frame.paragraphAlignment;
    }
    fontId = frame.fontId;
    fontName = frame.fontName;
  };

  while (tokenIndex < tokens.length) {
    const token = tokens[tokenIndex];
    tokenIndex += 1;
    if (!token) break;
    const currentFrame = groupFrames.at(-1);
    const destination = currentFrame?.destination ?? "root";

    if (token.kind === "group-start") {
      handleGroupStart(token.position);
      continue;
    }
    if (token.kind === "group-end") {
      handleGroupEnd();
      continue;
    }
    if (currentFrame?.skip) {
      if (token.kind === "text" && unicodeSkip > 0) unicodeSkip -= 1;
      continue;
    }

    if (token.kind === "control") {
      if (!currentFrame) {
        applyControl(token, "root");
        continue;
      }
      if (token.word === "rtf" && groupFrames.length === 1 && !currentFrame.decided) {
        currentFrame.decided = true;
        continue;
      }
      if (!currentFrame.decided) {
        currentFrame.decided = true;
        if (token.word === "rtf") continue;
        currentFrame.starred = token.symbol === "*";
        if (currentFrame.starred) {
          currentFrame.skip = true;
          continue;
        }
        const ignored = token.word !== undefined && IGNORED_DESTINATIONS.has(token.word);
        if (ignored) {
          currentFrame.skip = true;
          continue;
        }
        if (token.word === "fonttbl") {
          currentFrame.destination = "font-table";
          continue;
        }
        if (token.word === "colortbl") {
          currentFrame.destination = "color-table";
          continue;
        }
        if (token.word !== undefined && !CONTENT_GROUP_CONTROLS.has(token.word)) {
          currentFrame.skip = true;
          continue;
        }
      } else if (token.symbol === "*") {
        continue;
      }
      applyControl(token, destination);
      continue;
    }

    if (token.kind === "hex") {
      if (unicodeSkip > 0) {
        unicodeSkip -= 1;
        continue;
      }
      appendText(String.fromCharCode(Number.parseInt(token.value, 16)));
      continue;
    }

    if (token.kind === "text") {
      if (destination === "font-table" && fontId !== undefined && fontName !== undefined) {
        fontName += token.value;
        if (token.value.includes(";")) {
          const name = fontName.replace(/;.*$/u, "").trim();
          if (fontId !== undefined && name) fonts.set(fontId, name);
          fontName = "";
        }
      } else if (destination === "color-table" && token.value.includes(";")) {
        pushColor();
      } else if (unicodeSkip > 0) {
        const fallback = token.value.slice(unicodeSkip);
        unicodeSkip = 0;
        if (fallback) appendText(fallback);
      } else {
        appendText(token.value);
      }
      continue;
    }
  }

  groupFrames.length = 0;
  const table = activeTable();
  if (table) closeRow(table);
  flushParagraph();

  return { sections: [{ children: blocks }] };
}
