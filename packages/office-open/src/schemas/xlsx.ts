import { z } from "zod";

const anyObject = z.record(z.string(), z.unknown()).describe("JSON object");

export const xlsxSchema = z
  .object({
    worksheets: z
      .array(
        z
          .object({
            name: z.string().optional().describe("Sheet display name"),
            children: z
              .array(
                z
                  .object({
                    cells: z
                      .array(anyObject.or(z.union([z.string(), z.number(), z.boolean()])))
                      .optional(),
                    height: z.number().optional(),
                    hidden: z.boolean().optional(),
                    rowNumber: z.number().optional(),
                  })
                  .passthrough(),
              )
              .optional()
              .describe('Worksheet rows, each with "cells" array. NOTE: use "children" not "rows"'),
            columns: z.array(anyObject).optional(),
            mergeCells: z.array(anyObject).optional(),
            freezePanes: anyObject.optional(),
            autoFilter: z.string().optional(),
            images: z.array(anyObject).optional(),
            charts: z.array(anyObject).optional(),
            dataValidations: z.array(anyObject).optional(),
            conditionalFormats: z.array(anyObject).optional(),
          })
          .passthrough(),
      )
      .describe("Workbook worksheets (required)"),
    title: z.string().optional(),
    subject: z.string().optional(),
    creator: z.string().optional(),
    keywords: z.string().optional(),
    description: z.string().optional(),
    lastModifiedBy: z.string().optional(),
    revision: z.number().optional(),
  })
  .passthrough();
