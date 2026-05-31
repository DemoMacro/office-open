import { z } from "zod";

const anyObject = z.record(z.string(), z.unknown()).describe("JSON object");

const slideChild = z.union([
  z.object({ shape: anyObject }).passthrough(),
  z.object({ picture: anyObject }).passthrough(),
  z.object({ table: anyObject }).passthrough(),
  z.object({ chart: anyObject }).passthrough(),
  z.object({ line: anyObject }).passthrough(),
  z.object({ connector: anyObject }).passthrough(),
  z.object({ video: anyObject }).passthrough(),
  z.object({ audio: anyObject }).passthrough(),
  z.object({ group: anyObject }).passthrough(),
  z.object({ smartart: anyObject }).passthrough(),
  anyObject,
]);

const slideComment = z
  .object({
    author: z.string(),
    text: z.string(),
    x: z.number(),
    y: z.number(),
    initials: z.string().optional(),
    date: z.string().optional(),
  })
  .passthrough();

export const pptxSchema = z
  .object({
    slides: z
      .array(
        z
          .object({
            children: z
              .array(slideChild.or(z.string()))
              .describe(
                "Slide content: shapes ({ shape: {...} }), pictures ({ picture: {...} }), tables, charts, groups, lines, etc.",
              ),
            background: anyObject.optional(),
            transition: anyObject.optional(),
            notes: z.string().optional(),
            layout: z
              .enum([
                "blank",
                "title",
                "tx",
                "twoColTx",
                "titleOnly",
                "obj",
                "secHead",
                "chart",
                "tbl",
                "picTx",
                "twoObj",
                "twoTxTwoObj",
                "objTx",
                "vertTx",
                "vertTitleAndTx",
              ])
              .optional(),
            master: z.string().optional(),
            comments: z.array(slideComment).optional(),
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
    show: z
      .object({
        loop: z.boolean().optional(),
        kiosk: z.boolean().optional(),
        showNarration: z.boolean().optional(),
        useTimings: z.boolean().optional(),
      })
      .passthrough()
      .optional(),
    title: z.string().optional(),
    subject: z.string().optional(),
    creator: z.string().optional(),
    keywords: z.string().optional(),
    description: z.string().optional(),
    lastModifiedBy: z.string().optional(),
    revision: z.number().optional(),
  })
  .passthrough();
