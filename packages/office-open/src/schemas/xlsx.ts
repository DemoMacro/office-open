import { z } from "zod";

const anyObject = z.record(z.string(), z.unknown()).describe("JSON object");

const cellValue = z.union([z.string(), z.number(), z.boolean(), z.null()]);

const cellOptions = z.union([
  z
    .object({
      value: cellValue.optional(),
      reference: z.string().optional(),
      styleIndex: z.number().optional(),
      style: anyObject.optional(),
      formula: anyObject.optional(),
    })
    .passthrough(),
  cellValue,
]);

const columnOptions = z
  .object({
    min: z.number(),
    max: z.number(),
    width: z.number().optional(),
    hidden: z.boolean().optional(),
    customWidth: z.boolean().optional(),
  })
  .passthrough();

const mergeCellOptions = z
  .object({
    from: z.object({ row: z.number(), col: z.number() }),
    to: z.object({ row: z.number(), col: z.number() }),
  })
  .passthrough();

const freezePaneOptions = z
  .object({
    row: z.number().optional(),
    col: z.number().optional(),
  })
  .passthrough();

const dataValidationOptions = z
  .object({
    sqref: z.string(),
    type: z
      .enum(["none", "whole", "decimal", "list", "date", "time", "textLength", "custom"])
      .optional(),
    operator: z
      .enum([
        "between",
        "notBetween",
        "equal",
        "notEqual",
        "greaterThan",
        "lessThan",
        "greaterThanOrEqual",
        "lessThanOrEqual",
      ])
      .optional(),
    formula1: z.string().optional(),
    formula2: z.string().optional(),
    allowBlank: z.boolean().optional(),
    showErrorMessage: z.boolean().optional(),
    errorTitle: z.string().optional(),
    error: z.string().optional(),
    showInputMessage: z.boolean().optional(),
    promptTitle: z.string().optional(),
    prompt: z.string().optional(),
  })
  .passthrough();

const conditionalFormatOptions = z
  .object({
    sqref: z.string(),
    rules: z.array(
      z
        .object({
          type: z.enum(["cellIs", "containsText", "expression", "top10", "aboveAverage"]),
          operator: z
            .enum([
              "lessThan",
              "lessThanOrEqual",
              "equal",
              "notEqual",
              "greaterThanOrEqual",
              "greaterThan",
              "between",
              "notBetween",
              "containsText",
              "notContains",
              "beginsWith",
              "endsWith",
            ])
            .optional(),
          formulas: z.array(z.string()).optional(),
          priority: z.number().optional(),
          dxfId: z.number().optional(),
        })
        .passthrough(),
    ),
  })
  .passthrough();

export const xlsxSchema = z
  .object({
    worksheets: z
      .array(
        z
          .object({
            name: z.string().optional().describe("Sheet display name"),
            rows: z
              .array(
                z
                  .object({
                    cells: z.array(cellOptions).optional(),
                    height: z.number().optional(),
                    hidden: z.boolean().optional(),
                    rowNumber: z.number().optional(),
                  })
                  .passthrough(),
              )
              .optional()
              .describe('Worksheet rows, each with "cells" array'),
            columns: z.array(columnOptions).optional(),
            mergeCells: z.array(mergeCellOptions).optional(),
            freezePanes: freezePaneOptions.optional(),
            autoFilter: z.string().optional(),
            images: z.array(anyObject).optional(),
            charts: z.array(anyObject).optional(),
            dataValidations: z.array(dataValidationOptions).optional(),
            conditionalFormats: z.array(conditionalFormatOptions).optional(),
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
