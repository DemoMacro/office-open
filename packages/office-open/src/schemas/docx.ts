import { z } from "zod";

const anyObject = z.record(z.string(), z.unknown()).describe("JSON object");

export const docxSchema = z
  .object({
    sections: z
      .array(
        z
          .object({
            properties: anyObject.optional(),
            children: z
              .array(anyObject.or(z.string()))
              .describe(
                "Section content: paragraphs ({ paragraph: {...} }), tables ({ table: {...} }), images, TOCs, etc.",
              ),
            headers: anyObject.optional(),
            footers: anyObject.optional(),
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
    characterSpacingControl: z.string().optional(),
    view: z.string().optional(),
    zoom: anyObject.optional(),
    writeProtection: anyObject.optional(),
    displayBackgroundShape: z.boolean().optional(),
    embedTrueTypeFonts: z.boolean().optional(),
    embedSystemFonts: z.boolean().optional(),
    saveSubsetFonts: z.boolean().optional(),
    docVars: z.array(anyObject).optional(),
    colorSchemeMapping: z.unknown().optional(),
  })
  .passthrough();
