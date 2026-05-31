import { z } from "zod";

const anyObject = z.record(z.string(), z.unknown()).describe("JSON object");

export const pptxSchema = z
  .object({
    slides: z
      .array(
        z
          .object({
            children: z
              .array(anyObject.or(z.string()))
              .describe(
                "Slide content: shapes ({ shape: {...} }), pictures ({ picture: {...} }), tables, charts, groups, lines, etc.",
              ),
            background: anyObject.optional(),
            transition: anyObject.optional(),
            notes: z.string().optional(),
            layout: z.string().optional(),
            master: z.string().optional(),
            comments: z.array(anyObject).optional(),
            headerFooter: anyObject.optional(),
          })
          .passthrough(),
      )
      .describe("Presentation slides (required)"),
    size: z
      .union([z.enum(["16:9", "4:3"]), z.object({ width: z.number(), height: z.number() })])
      .optional()
      .describe('Slide size: "16:9", "4:3", or { width, height }'),
    masters: z.array(anyObject).optional().describe("Slide master definitions"),
    show: anyObject.optional(),
    title: z.string().optional(),
    subject: z.string().optional(),
    creator: z.string().optional(),
    keywords: z.string().optional(),
    description: z.string().optional(),
    lastModifiedBy: z.string().optional(),
    revision: z.number().optional(),
  })
  .passthrough();
