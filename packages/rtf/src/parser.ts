import type {
  DocumentOptions,
  FootnoteOptions,
  NumberingOptions,
  ParagraphOptions,
  SectionChild,
  SectionOptions,
  RunOptions,
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
  "category",
  "company",
  "datafield",
  "datastore",
  "doccomm",
  "falttext",
  "filetbl",
  "footerf",
  "footerl",
  "footerr",
  "headerf",
  "headerl",
  "headerr",
  "hlinkbase",
  "hlink",
  "info",
  "keywords",
  "listoverridetable",
  "manager",
  "operator",
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
  const fonts = new Map<number, string>();
  const colors: (string | null)[] = [];
  const tables: TableDraft[] = [];
  const groupFrames: GroupFrame[] = [];
  const sections: SectionOptions[] = [{ children: [] }];
  let blocks: SectionChild[] = sections[0]!.children;
  const footnotes: FootnoteOptions[] = [];
  const numberingReferences = new Set<string>();
  const listFormats = new Map<string, "bullet" | "decimal">();
  const headerBlocks = new Map<"default" | "first", SectionChild[]>();
  const footerBlocks = new Map<"default" | "first", SectionChild[]>();
  let objectType: string | undefined;
  let bookmarkId = 1;
  const bookmarkIds = new Map<string, number>();
  let listReference = "rtf-list-1";
  let pendingList: { reference: string; level?: number } | undefined;
  let pendingShading: NonNullable<ParagraphOptions["shading"]> | undefined;
  let pendingBorder: NonNullable<ParagraphOptions["border"]> | undefined;

  let format: RunFormat = {};
  let alignment: ParagraphOptions["alignment"];
  let paragraph: ParagraphDraft | undefined;
  let unicodeByteCount = 1;
  let unicodeSkip = 0;
  let fontId: number | undefined;
  let fontName: string | undefined;
  let color: { red?: number; green?: number; blue?: number } = {};
  let tokenIndex = 0;

  const consumeGroup = (start: number): number => {
    let depth = 1;
    let index = start;
    while (index < tokens.length && depth > 0) {
      const token = tokens[index];
      if (token?.kind === "group-start") depth += 1;
      else if (token?.kind === "group-end") depth -= 1;
      index += 1;
    }
    return depth === 0 ? index : tokens.length;
  };

  const decodeDestinationText = (start: number, end: number): string[] => {
    const lines: string[] = [];
    let line = "";
    let skip = 0;
    let byteCount = 1;
    for (let index = start; index < end; index += 1) {
      const token = tokens[index];
      if (!token) continue;
      if (token.kind === "group-start" || token.kind === "group-end") continue;
      if (token.kind === "hex") {
        if (skip > 0) skip -= 1;
        else line += String.fromCharCode(Number.parseInt(token.value, 16));
        continue;
      }
      if (token.kind === "text") {
        if (skip > 0) {
          line += token.value.slice(skip);
          skip = 0;
        } else line += token.value;
        continue;
      }
      if (token.kind !== "control") continue;
      if (token.word === "par") {
        lines.push(line);
        line = "";
      } else if (token.word === "line" || token.word === "tab") {
        line += token.word === "tab" ? "\t" : "\n";
      } else if (token.word === "u" && token.param !== undefined) {
        const codePoint = token.param < 0 ? token.param + 65536 : token.param;
        if (codePoint >= 0 && codePoint <= 0x10ffff) line += String.fromCodePoint(codePoint);
        skip = byteCount;
      } else if (token.word === "uc") {
        byteCount = token.param ?? 1;
      }
    }
    lines.push(line);
    return lines.map((value) => value.trim()).filter((value) => value.length > 0);
  };

  const blocksFromText = (start: number, end: number): SectionChild[] =>
    decodeDestinationText(start, end).map((text) => ({ paragraph: { children: [{ text }] } }));

  const readPicture = (start: number, end: number) => {
    let type: "png" | "jpg" | "wmf" | "emf" = "png";
    let width = 0;
    let height = 0;
    let hex = "";
    for (let index = start; index < end; index += 1) {
      const token = tokens[index];
      if (token?.kind === "control") {
        if (token.word === "pngblip") type = "png";
        else if (token.word === "jpegblip") type = "jpg";
        else if (token.word === "emfblip") type = "emf";
        else if (token.word === "wmetafile") type = "wmf";
        else if ((token.word === "picw" || token.word === "picwgoal") && token.param !== undefined)
          width = token.param * 635;
        else if ((token.word === "pich" || token.word === "pichgoal") && token.param !== undefined)
          height = token.param * 635;
      } else if (token?.kind === "hex") hex += token.value;
      else if (token?.kind === "text") hex += token.value.replace(/[^0-9a-fA-F]/g, "");
    }
    if (!hex || width <= 0 || height <= 0 || hex.length % 2 !== 0) return undefined;
    try {
      const data = new Uint8Array((hex.length / 2) | 0);
      for (let index = 0; index < data.length; index += 1)
        data[index] = Number.parseInt(hex.slice(index * 2, index * 2 + 2), 16);
      if (data.length === 0) return undefined;
      return { type, data, width, height };
    } catch {
      return undefined;
    }
  };

  const readField = (start: number, end: number) => {
    let instruction = "";
    let result = "";
    let target: "instruction" | "result" | undefined;
    let depth = 0;
    for (let index = start; index < end; index += 1) {
      const token = tokens[index];
      if (!token) continue;
      if (token.kind === "group-start") depth += 1;
      else if (token.kind === "group-end") depth -= 1;
      if (token.kind === "control") {
        if (token.word === "fldinst") target = "instruction";
        else if (token.word === "fldrslt") target = "result";
        continue;
      }
      if (target === "instruction" && token.kind === "text") instruction += token.value;
      if (target === "result" && (token.kind === "text" || token.kind === "hex")) {
        result +=
          token.kind === "text"
            ? token.value
            : String.fromCharCode(Number.parseInt(token.value, 16));
      }
    }
    instruction = instruction.trim();
    result = result.trim();
    if (!instruction) return undefined;
    if (/^HYPERLINK\s+"([^"]+)"/iu.test(instruction)) {
      const url = RegExp.$1;
      return { hyperlink: { url, children: result ? [result] : [] } };
    }
    if (/^PAGEREF\s+(\S+)/iu.test(instruction))
      return { pageReference: { bookmarkId: RegExp.$1, hyperlink: true } };
    return { complexField: { instruction, ...(result ? { result } : {}) } };
  };

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
    const draft = ensureParagraph();
    if (pendingList) draft.options.numbering = { ...pendingList };
    if (pendingShading) draft.options.shading = { ...pendingShading };
    if (pendingBorder) draft.options.border = { ...pendingBorder };
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
        pendingList = undefined;
        pendingShading = undefined;
        pendingBorder = undefined;
        if (paragraph) paragraph.options.numbering = undefined;
        return;
      case "sectd":
        if (activeTable()) return;
        flushParagraph();
        blocks = [];
        return;
      case "sect": {
        if (activeTable()) return;
        flushParagraph(true);
        const next: SectionOptions = { children: [] };
        sections.push(next);
        blocks = next.children;
        return;
      }
      case "pgwsxn":
      case "pghsxn":
        if (activeTable()) return;
        sections[0]!.properties ??= {};
        sections[0]!.properties.pageSize ??= {};
        if (sections[0]!.properties.pageSize !== false && token.param !== undefined) {
          if (word === "pgwsxn") sections[0]!.properties.pageSize.width = token.param;
          else sections[0]!.properties.pageSize.height = token.param;
        }
        return;
      case "marglsxn":
      case "margsxn":
      case "margtsxn":
      case "margbsxn":
        if (activeTable()) return;
        sections[0]!.properties ??= {};
        sections[0]!.properties.pageMargin ??= {};
        if (sections[0]!.properties.pageMargin !== false && token.param !== undefined) {
          const margin = sections[0]!.properties.pageMargin;
          if (word === "marglsxn") margin.left = token.param;
          else if (word === "margsxn") margin.right = token.param;
          else if (word === "margtsxn") margin.top = token.param;
          else margin.bottom = token.param;
        }
        return;
      case "ls":
        if (token.param !== undefined) {
          listReference = `rtf-list-${token.param}`;
          numberingReferences.add(listReference);
          pendingList = { reference: listReference };
        }
        return;
      case "ilvl":
        if (pendingList && token.param !== undefined) pendingList.level = token.param;
        const numbering = paragraph?.options.numbering;
        if (numbering && "reference" in numbering && token.param !== undefined)
          numbering.level = token.param;
        return;
      case "shading":
      case "cbpat": {
        const draft = ensureParagraph();
        const fill = word === "cbpat" ? colors[token.param ?? 0] : "auto";
        pendingShading = { type: "clear", fill: typeof fill === "string" ? fill : "auto" };
        draft.options.shading = { ...pendingShading };
        return;
      }
      case "brdrb":
      case "brdrt":
      case "brdrl":
      case "brdrr": {
        const border = {
          style: "single",
          ...(token.param !== undefined ? { size: token.param } : {}),
        };
        pendingBorder ??= {};
        if (word === "brdrb") pendingBorder.bottom = border;
        else if (word === "brdrt") pendingBorder.top = border;
        else if (word === "brdrl") pendingBorder.left = border;
        else pendingBorder.right = border;
        const draft = ensureParagraph();
        draft.options.border = { ...pendingBorder };
        return;
      }
      case "objclass":
        objectType = undefined;
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
          if (token.word === "bkmkstart" || token.word === "bkmkend") {
            currentFrame.skip = false;
            const end = consumeGroup(tokenIndex);
            const name = decodeDestinationText(tokenIndex, end)[0] ?? "";
            if (name) {
              if (token.word === "bkmkstart") {
                const id = bookmarkId++;
                bookmarkIds.set(name, id);
                appendInline({ bookmarkStart: { id, name } });
              } else {
                const id = bookmarkIds.get(name) ?? bookmarkId++;
                bookmarkIds.delete(name);
                appendInline({ bookmarkEnd: { id } });
              }
            }
            groupFrames.pop();
            tokenIndex = end;
            continue;
          }
          currentFrame.skip = true;
          continue;
        }
        if (token.word === "bkmkstart" || token.word === "bkmkend") {
          const end = consumeGroup(tokenIndex);
          const name = decodeDestinationText(tokenIndex, end)[0] ?? "";
          if (name) {
            if (token.word === "bkmkstart") {
              const id = bookmarkId++;
              bookmarkIds.set(name, id);
              appendInline({ bookmarkStart: { id, name } });
            } else {
              const id = bookmarkIds.get(name) ?? bookmarkId++;
              bookmarkIds.delete(name);
              appendInline({ bookmarkEnd: { id } });
            }
          }
          groupFrames.pop();
          tokenIndex = end;
          continue;
        }
        if (
          token.word === "pict" ||
          token.word === "field" ||
          token.word?.startsWith("header") ||
          token.word?.startsWith("footer") ||
          token.word === "footnote" ||
          token.word === "objdata" ||
          token.word === "pn" ||
          token.word === "object"
        ) {
          try {
            const end = consumeGroup(tokenIndex);
            if (token.word === "pict") {
              const picture = readPicture(tokenIndex, end);
              if (picture)
                appendInline({
                  picture: {
                    type: picture.type,
                    data: picture.data,
                    transformation: { width: picture.width, height: picture.height },
                  },
                });
            } else if (token.word === "field") {
              const child = readField(tokenIndex, end);
              if (child) appendInline(child);
            } else if (token.word?.startsWith("header") || token.word?.startsWith("footer")) {
              const children = blocksFromText(tokenIndex, end);
              const target = token.word.endsWith("f") ? "first" : "default";
              const map = token.word?.startsWith("header") ? headerBlocks : footerBlocks;
              map.set(target, children);
            } else if (token.word === "footnote") {
              const children = blocksFromText(tokenIndex, end);
              if (children.length > 0) footnotes.push({ children });
            } else if (token.word === "objdata") {
              appendText(objectType ? `[Embedded object: ${objectType}]` : "[Embedded object]");
            } else if (token.word === "pn") {
              for (let index = tokenIndex; index < end; index += 1) {
                const nested = tokens[index];
                if (nested?.kind === "control") {
                  if (nested.word === "pndec") listFormats.set(listReference, "decimal");
                  if (nested.word === "pnbullet") listFormats.set(listReference, "bullet");
                }
              }
            } else if (token.word === "object") {
              let targetClass = false;
              let className = "";
              for (let index = tokenIndex; index < end; index += 1) {
                const nested = tokens[index];
                if (nested?.kind === "control") {
                  if (nested.word === "objclass") {
                    targetClass = true;
                    className = "";
                  } else targetClass = false;
                } else if (targetClass && nested?.kind === "text") className += nested.value;
              }
              if (className) objectType = className.trim();
              appendText(objectType ? `[Embedded object: ${objectType}]` : "[Embedded object]");
            }
            groupFrames.pop();
            tokenIndex = end;
          } catch {
            groupFrames.pop();
            currentFrame.skip = true;
          }
          continue;
        }
        const ignored = token.word !== undefined && IGNORED_DESTINATIONS.has(token.word);
        if (token.word === "ls") {
          currentFrame.skip = false;
          applyControl(token, destination);
          continue;
        }
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

  const first = sections[0];
  if (first) {
    const headers = headerBlocks.get("default");
    const firstHeaders = headerBlocks.get("first");
    const footers = footerBlocks.get("default");
    const firstFooters = footerBlocks.get("first");
    if (headers) first.headers = { default: headers };
    if (firstHeaders) first.headers = { ...first.headers, first: firstHeaders };
    if (footers) first.footers = { default: footers };
    if (firstFooters) first.footers = { ...first.footers, first: firstFooters };
  }

  const numbering: NumberingOptions | undefined =
    numberingReferences.size > 0
      ? {
          abstractNumberings: [...numberingReferences].map((reference) => ({
            reference,
            levels: [
              {
                level: 0,
                format: listFormats.get(reference) === "bullet" ? "bullet" : "decimal",
                text: listFormats.get(reference) === "bullet" ? "●" : "%1.",
              },
            ],
          })),
        }
      : undefined;

  return {
    sections,
    ...(numbering ? { numbering } : {}),
    ...(footnotes.length > 0 ? { footnotes } : {}),
  };
}
