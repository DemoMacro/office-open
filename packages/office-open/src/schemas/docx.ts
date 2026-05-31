import { z } from "zod";

const anyObject = z.record(z.string(), z.unknown()).describe("JSON object");

const sectionChild = z.union([
  z.object({ paragraph: anyObject }).passthrough(),
  z.object({ table: anyObject }).passthrough(),
  z.object({ toc: anyObject }).passthrough(),
  z.object({ textbox: anyObject }).passthrough(),
  z.object({ sdt: anyObject }).passthrough(),
  z.object({ altChunk: anyObject }).passthrough(),
  z.object({ subDoc: anyObject }).passthrough(),
  anyObject,
]);

const headerFooterGroup = z
  .object({
    default: anyObject.or(z.array(sectionChild.or(z.string()))).optional(),
    first: anyObject.or(z.array(sectionChild.or(z.string()))).optional(),
    even: anyObject.or(z.array(sectionChild.or(z.string()))).optional(),
  })
  .passthrough();

export const docxSchema = z
  .object({
    sections: z
      .array(
        z
          .object({
            properties: anyObject.optional(),
            children: z
              .array(sectionChild.or(z.string()))
              .describe(
                "Section content: paragraphs ({ paragraph: {...} }), tables ({ table: {...} }), images, TOCs, etc.",
              ),
            headers: headerFooterGroup.optional(),
            footers: headerFooterGroup.optional(),
          })
          .passthrough(),
      )
      .describe("Document sections (required)"),
    title: z.string().optional(),
    subject: z.string().optional(),
    creator: z.string().optional(),
    keywords: z.string().optional(),
    description: z.string().optional(),
    lastModifiedBy: z.string().optional(),
    revision: z.number().optional(),
    externalStyles: z.string().optional(),
    styles: anyObject.optional(),
    numbering: anyObject.optional(),
    comments: anyObject.optional(),
    bibliography: anyObject.optional(),
    footnotes: z.record(z.string(), anyObject).optional(),
    endnotes: z.record(z.string(), anyObject).optional(),
    background: anyObject.optional(),
    features: anyObject.optional(),
    compatabilityModeVersion: z.number().optional(),
    compatibility: anyObject.optional(),
    customProperties: z.array(anyObject).optional(),
    evenAndOddHeaderAndFooters: z.boolean().optional(),
    defaultTabStop: z.number().optional(),
    fonts: z.array(anyObject).optional(),
    hyphenation: anyObject.optional(),
    characterSpacingControl: z.enum(["compressPunctuation", "doNotCompress"]).optional(),
    view: z.enum(["none", "print", "outline", "masterPages", "normal", "web"]).optional(),
    zoom: z
      .object({
        percent: z.number().optional(),
        val: z.enum(["none", "fullPage", "bestFit", "textFit"]).optional(),
      })
      .passthrough()
      .optional(),
    writeProtection: anyObject.optional(),
    displayBackgroundShape: z.boolean().optional(),
    embedTrueTypeFonts: z.boolean().optional(),
    embedSystemFonts: z.boolean().optional(),
    saveSubsetFonts: z.boolean().optional(),
    docVars: z.array(z.object({ name: z.string(), val: z.string() }).passthrough()).optional(),
    colorSchemeMapping: anyObject.optional(),
  })
  .passthrough();
