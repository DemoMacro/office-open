import { Relationships, type XmlifyedFile, replaceNumberingPlaceholders } from "@office-open/core";
import { escapeXml } from "@office-open/xml";

import { stringifyBodyChild, type BodyContext } from "../body";
import type { DocxWriteContext } from "../context";
import { stringifyWpsTextBoxChild } from "../parts/drawing/descriptor";
import type { TextBoxPartOptions } from "../parts/drawing/inline/graphic/graphic-data/wps/wps-shape";
import { XML_DECL, registerPartMedia, resolvePartCharts, resolvePartMedia } from "./shared";

function collectTextBoxParts(value: unknown, parts: TextBoxPartOptions[], seen: Set<object>): void {
  if (!value || typeof value !== "object" || seen.has(value)) return;
  seen.add(value);
  if (Array.isArray(value)) {
    for (const child of value) collectTextBoxParts(child, parts, seen);
    return;
  }
  const record = value as Record<string, unknown>;
  const part = record.textBoxPart;
  if (
    typeof part === "object" &&
    part !== null &&
    typeof (part as TextBoxPartOptions).path === "string" &&
    Array.isArray((part as TextBoxPartOptions).children)
  ) {
    parts.push(part as TextBoxPartOptions);
  }
  for (const child of Object.values(record)) collectTextBoxParts(child, parts, seen);
}

interface CompiledTextBoxPart {
  part: XmlifyedFile;
  relationships?: XmlifyedFile;
}

export function compileTextBoxParts(ctx: DocxWriteContext): CompiledTextBoxPart[] {
  const options: TextBoxPartOptions[] = [];
  collectTextBoxParts(ctx._options, options, new Set());
  const seenPaths = new Set<string>();
  const compiled: CompiledTextBoxPart[] = [];
  for (const optionsPart of options) {
    if (seenPaths.has(optionsPart.path)) continue;
    seenPaths.add(optionsPart.path);
    const relationships = new Relationships(optionsPart.path);
    const bodyCtx: BodyContext = {
      fileData: ctx,
      file: ctx,
      viewWrapper: { relationships, partName: optionsPart.path },
      reproducible: ctx.reproducible,
      addRelationship: (type, target, mode) => ctx.addRelationship(type, target, mode),
      addMedia: (data, type, fileName) => ctx.addMedia(data, type, fileName),
      addHyperlink: (key, target) => ctx.addHyperlink(key, target),
      stringifyChild: undefined as unknown as BodyContext["stringifyChild"],
    };
    bodyCtx.stringifyChild = (child) => stringifyBodyChild(child, bodyCtx);
    const children = (optionsPart.children ?? [])
      .map((child) => stringifyWpsTextBoxChild(child, bodyCtx))
      .join("");
    const ignorable = optionsPart.ignorable
      ? ` mc:Ignorable="${escapeXml(optionsPart.ignorable)}"`
      : "";
    const source = `${XML_DECL}<w14:txbx xmlns:w14="http://schemas.microsoft.com/office/word/2008/9/12/wordml" xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"${ignorable}>${children}</w14:txbx>`;
    const relationshipCount = relationships.nextRelationshipId;
    const media = resolvePartMedia(source, ctx, relationshipCount);
    registerPartMedia(relationships, ctx, media);
    const xml = replaceNumberingPlaceholders(
      resolvePartCharts(
        media.xml,
        ctx,
        relationships,
        media.embeddingOffset + media.embeddingRefs.length,
        optionsPart.path,
      ),
      ctx.numbering.concreteNumbering,
    );
    const result: CompiledTextBoxPart = {
      part: { data: xml, path: optionsPart.path },
    };
    if (relationships.relationshipCount > 0) {
      result.relationships = {
        data: XML_DECL + relationships.serialize(),
        path: `word/_rels/${optionsPart.path.slice("word/".length)}.rels`,
      };
    }
    compiled.push(result);
  }
  return compiled;
}
