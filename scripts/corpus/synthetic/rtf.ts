import { generateDocument, parseDocument } from "../../../packages/docx/dist/index.mjs";
import { parseDocument as parseRtfDocument } from "../../../packages/rtf/dist/index.mjs";
import { assert, assertEqual } from "./support";

async function projectRtf(source: string, part: string): Promise<void> {
  const output = await generateDocument(parseRtfDocument(source), { type: "uint8array" });
  await parseDocument(output);
}

const BINARY_SOURCE = String.raw`{\rtf1{\object{\objclass Excel.Sheet.8}{\objdata\bin4 ABCD}}}`;

export async function rtfBinaryToken(): Promise<void> {
  const part = "RTF objdata";
  const options = parseRtfDocument(BINARY_SOURCE);
  const paragraph = options.sections[0]?.children?.[0]?.paragraph;
  const group = paragraph && "children" in paragraph ? paragraph.children?.[0] : undefined;
  const object = group && "children" in group ? group.children?.[0] : undefined;
  assert(object && "object" in object, part, "embedded object projection");
  assertEqual(object.object?.embed?.data?.join(","), "65,66,67,68", part, "binary token bytes");
  await projectRtf(BINARY_SOURCE, part);
}

const STYLESHEET_SOURCE = String.raw`{\rtf1{\colortbl;\red255\green0\blue0;}{\stylesheet{\s7 Base;\sbasedon7\snext7\shidden\spriority9\sqformat\ql\li120\sb240\brdrb\brdrw15\cbpat1}{\cs8 Link;\b\i}}\s7 Body}`;

export async function rtfStylesheet(): Promise<void> {
  const part = "RTF stylesheet";
  const options = parseRtfDocument(STYLESHEET_SOURCE);
  assertEqual(options.styles?.paragraphStyles?.[0]?.id, "rtf-style-7", part, "paragraph style id");
  assertEqual(options.styles?.paragraphStyles?.[0]?.uiPriority, 9, part, "style priority");
  assertEqual(options.styles?.characterStyles?.[0]?.run?.bold, true, part, "character style run");
  await projectRtf(STYLESHEET_SOURCE, part);
}

const LISTTABLE_SOURCE = String.raw`{\rtf1{\listtable{\list\listtemplateid9\listhybrid{\listlevel\levelnfc2\leveljc1\levelstartat3\levelindent720{\leveltext\'02\'00.;}{\levelnumbers;}}{\listname rtf-list-template-9;}\listid77}}{\listoverridetable{\listoverride\listid77\listoverridecount0\ls4}}\ls4\ilvl0\par Item}`;

export async function rtfListtable(): Promise<void> {
  const part = "RTF listtable";
  const options = parseRtfDocument(LISTTABLE_SOURCE);
  assertEqual(
    options.numbering?.abstractNumberings?.[0]?.reference,
    "rtf-list-template-9",
    part,
    "list template reference",
  );
  assertEqual(
    options.numbering?.abstractNumberings?.[0]?.levels?.[0]?.format,
    "lowerRoman",
    part,
    "list level format",
  );
  await projectRtf(LISTTABLE_SOURCE, part);
}
