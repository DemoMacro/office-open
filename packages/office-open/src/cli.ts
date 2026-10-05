import { readFile, writeFile } from "node:fs/promises";

import { OOXML_PACKAGE_FORMATS } from "@office-open/core";
import { defineCommand, runMain } from "citty";

import { detectOffice, type OfficeFormatInfo } from "./formats";
import { generateOffice, type OfficeGenerateFormat, type OfficeOptionsFor } from "./generate";
import { parseOffice, type ParsedOffice } from "./parse";
import {
  SCHEMA_ENTRIES,
  UnknownDefinitionError,
  renderSliceTypeText,
  sliceDocumentSchema,
  validateDocumentInput,
} from "./schemas";
import { SCHEMAS, type DocumentType } from "./schemas/schemas";

type OoxmlGenerateFormat = keyof typeof OOXML_PACKAGE_FORMATS;

type OfficeGenerateCommandFormat = OoxmlGenerateFormat | "odt" | "ods" | "odp" | "rtf";

type GenerateFormat = OfficeGenerateCommandFormat;

const FORMATS = [
  ...Object.keys(OOXML_PACKAGE_FORMATS),
  "odt",
  "ods",
  "odp",
  "rtf",
] as OfficeGenerateCommandFormat[];

const SCHEMA_TYPES = {
  wordprocessing: "docx",
  spreadsheet: "xlsx",
  presentation: "pptx",
  odt: "docx",
  ods: "xlsx",
  odp: "pptx",
  rtf: "docx",
} as const;

function schemaTypeOf(format: OfficeGenerateCommandFormat): DocumentType {
  if (format === "docx" || format === "pptx" || format === "xlsx") return format;
  return SCHEMA_TYPES[format as keyof typeof SCHEMA_TYPES];
}

/** Parse and validate a generate-format positional (citty positionals cannot be enums). */
function parseGenerateFormat(raw: string | undefined): OfficeGenerateCommandFormat {
  if ((FORMATS as readonly string[]).includes(raw ?? "")) {
    return raw as OfficeGenerateCommandFormat as GenerateFormat;
  }
  console.error(`Unknown format "${raw ?? ""}" — expected one of: ${FORMATS.join(", ")}`);
  globalThis.process.exitCode = 1;
  throw new Error("invalid format");
}

/** Parse and validate a schema-family positional. */
function parseSchemaFormat(raw: string | undefined): DocumentType {
  if (raw === "docx" || raw === "pptx" || raw === "xlsx") return raw;
  const format = parseGenerateFormat(raw);
  return schemaTypeOf(format);
}

async function readJsonInput(input: string): Promise<Record<string, unknown>> {
  const trimmed = input.trim();
  if (trimmed.startsWith("{") || trimmed.startsWith("[")) {
    return JSON.parse(trimmed) as Record<string, unknown>;
  }
  return JSON.parse(await readFile(input, "utf-8")) as Record<string, unknown>;
}

async function detectOfficePath(path: string): Promise<OfficeFormatInfo> {
  return detectOffice(new Uint8Array(await readFile(path)));
}

async function parseOfficePath(path: string): Promise<ParsedOffice> {
  return parseOffice(new Uint8Array(await readFile(path)));
}

async function generateOfficePath<Format extends OfficeGenerateFormat>(
  path: string,
  format: Format,
  options: OfficeOptionsFor[Format],
): Promise<void> {
  const contents = await generateOffice(format, options, "nodebuffer");
  await writeFile(path, contents);
}

function createConvertCommand(type: string, defaultExt: string) {
  return defineCommand({
    meta: {
      name: type,
      description: `Generate a .${defaultExt} file from JSON`,
    },
    args: {
      input: {
        type: "positional",
        description: "JSON string or path to JSON file",
        required: true,
      },
      output: {
        type: "positional",
        description: `Output file path (default: output.${defaultExt})`,
        required: false,
      },
      "input-file": {
        type: "string",
        description: "Read JSON from file (alternative to positional input)",
        alias: ["i"],
      },
      "output-file": {
        type: "string",
        description: "Output file path (alternative to positional output)",
        alias: ["o"],
      },
    },
    async run({ args }) {
      const jsonInput = (args.input ?? args["input-file"]) as string;
      const outputPath = (args.output ?? args["output-file"] ?? `output.${defaultExt}`) as string;
      const generateType = type as GenerateFormat;
      const docType = schemaTypeOf(generateType);

      try {
        const docOptions = await readJsonInput(jsonInput);
        const validated = validateDocumentInput(docType, docOptions);
        await generateOfficePath(outputPath, generateType, validated);
        console.log(`Generated: ${outputPath}`);
      } catch (error) {
        // Expected user errors (bad JSON, schema violations) print as a single line;
        // rethrowing would make runMain dump them with a stack trace.
        console.error(`Error: ${(error as Error).message}`);
        globalThis.process.exitCode = 1;
      }
    },
  });
}

const schemaIndexCommand = defineCommand({
  meta: {
    name: "index",
    description: "List schema definitions (indexed lookup entries by default)",
  },
  args: {
    format: { type: "positional", description: "docx | pptx | xlsx", required: true },
    all: { type: "boolean", alias: "a", description: "List every definition name" },
    json: { type: "boolean", description: "Machine-readable output" },
  },
  run({ args }) {
    const format = parseSchemaFormat(args.format as string | undefined);
    const definitionCount = Object.keys(
      (SCHEMAS[format].definitions as Record<string, unknown>) ?? {},
    ).length;

    if (args.json) {
      console.log(
        JSON.stringify(
          args.all
            ? {
                format,
                definitionCount,
                definitions: Object.keys(SCHEMAS[format].definitions as Record<string, unknown>),
              }
            : { format, definitionCount, entries: SCHEMA_ENTRIES[format] },
          null,
          2,
        ),
      );
      return;
    }

    if (args.all) {
      for (const name of Object.keys(SCHEMAS[format].definitions as Record<string, unknown>)) {
        console.log(name);
      }
      return;
    }

    const entries = SCHEMA_ENTRIES[format];
    console.log(`${format}: ${entries.length} lookup entries of ${definitionCount} definitions`);
    console.log();
    let currentDomain = "";
    for (const entry of entries) {
      if (entry.domain !== currentDomain) {
        currentDomain = entry.domain;
        console.log(`${currentDomain}`);
      }
      console.log(`  ${entry.name.padEnd(44)} ${entry.summary}`);
    }
    console.log();
    console.log(`Slice a definition's fields (--json for the raw schema):`);
    console.log(`  office-open schema slice ${format} <Definition> [more...]`);
    console.log(`List every definition name:`);
    console.log(`  office-open schema index ${format} --all`);
  },
});

const schemaSliceCommand = defineCommand({
  meta: {
    name: "slice",
    description: "Print a definition slice as type definitions (--json for the raw JSON schema)",
  },
  args: {
    format: { type: "positional", description: "docx | pptx | xlsx", required: true },
    definitions: {
      type: "positional",
      description: "One or more definition names (variadic)",
      required: true,
    },
    json: { type: "boolean", description: "Emit the raw draft-07 JSON schema instead" },
  },
  run({ args }) {
    // citty does not type variadic positionals; args._ keeps every raw positional
    // (format first), so slice the tail off it instead of the typed args.
    const positional = args._ as string[];
    const format = parseSchemaFormat(positional[0]);
    const definitions = positional.slice(1);
    if (definitions.length === 0) {
      console.error("Provide at least one definition name (see `office-open schema index`).");
      globalThis.process.exitCode = 1;
      return;
    }
    try {
      const slice = sliceDocumentSchema(format, definitions);
      console.log(
        args.json
          ? JSON.stringify(slice, null, 2)
          : renderSliceTypeText(format, definitions, slice),
      );
    } catch (error) {
      if (error instanceof UnknownDefinitionError) {
        console.error(`${error.message}`);
        if (error.suggestions.length > 0) {
          console.error(`Closest: ${error.suggestions.join(", ")}`);
        }
        console.error(`List all names with: office-open schema index ${format} --all`);
      } else {
        throw error;
      }
      globalThis.process.exitCode = 1;
    }
  },
});

const schemaCommand = defineCommand({
  meta: {
    name: "schema",
    description: "Consult the JSON schemas: list lookup entries or slice definitions",
  },
  subCommands: { index: schemaIndexCommand, slice: schemaSliceCommand },
});

const detectCommand = defineCommand({
  meta: {
    name: "detect",
    description: "Detect the office format of a binary or text file",
  },
  args: {
    input: { type: "positional", description: "File path", required: true },
    json: { type: "boolean", description: "Machine-readable output" },
  },
  async run({ args }) {
    try {
      const info = await detectOfficePath(args.input as string);
      if (args.json) {
        console.log(JSON.stringify(info));
        return;
      }
      console.log(`Format: ${info.format}`);
      console.log(`Family: ${info.family}`);
      console.log(`Container: ${info.container}`);
    } catch (error) {
      console.error(`Error: ${(error as Error).message}`);
      globalThis.process.exitCode = 1;
    }
  },
});

const parseCommand = defineCommand({
  meta: {
    name: "parse",
    description: "Parse a supported office file into its Options JSON model",
  },
  args: {
    input: { type: "positional", description: "File path", required: true },
    output: {
      type: "string",
      description: "Write JSON to a file instead of stdout",
      alias: ["o"],
    },
  },
  async run({ args }) {
    try {
      const parsed = await parseOfficePath(args.input as string);
      const json = JSON.stringify(parsed, null, 2);
      if (args.output) {
        await writeFile(args.output as string, json);
        console.log(`Parsed: ${args.output}`);
        return;
      }
      console.log(json);
    } catch (error) {
      console.error(`Error: ${(error as Error).message}`);
      globalThis.process.exitCode = 1;
    }
  },
});

const convertCommands = Object.fromEntries(
  FORMATS.map((format) => [format, createConvertCommand(format, format)]),
) as Record<GenerateFormat, ReturnType<typeof createConvertCommand>>;

const mainCommand = defineCommand({
  meta: {
    name: "office-open",
    version: "0.14.6",
    description: "Generate OOXML and OpenDocument files from JSON",
  },
  subCommands: {
    ...convertCommands,
    detect: detectCommand,
    parse: parseCommand,
    schema: schemaCommand,
  },
  args: {
    type: {
      type: "enum",
      description: "File type to generate",
      options: [...FORMATS],
    },
  },
  async run() {
    // citty shows usage when no subcommand is matched
  },
});

void runMain(mainCommand);
