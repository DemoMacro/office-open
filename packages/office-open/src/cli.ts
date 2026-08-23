import { defineCommand, runMain } from "citty";

import { generateToFile, parseInput } from "./generate";
import {
  SCHEMA_ENTRIES,
  UnknownDefinitionError,
  renderSliceTypeText,
  sliceDocumentSchema,
  validateDocumentInput,
} from "./schemas";
import { SCHEMAS, type DocumentType } from "./schemas/schemas";

const FORMATS: readonly DocumentType[] = ["docx", "pptx", "xlsx"];

/** Parse and validate the format positional (citty positionals cannot be enums). */
function parseFormat(raw: string | undefined): DocumentType {
  if (raw && (FORMATS as readonly string[]).includes(raw)) return raw as DocumentType;
  console.error(`Unknown format "${raw ?? ""}" — expected one of: ${FORMATS.join(", ")}`);
  globalThis.process.exitCode = 1;
  throw new Error("invalid format");
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
      const docType = type as "docx" | "pptx" | "xlsx";

      try {
        const docOptions = await parseInput(jsonInput);
        const validated = validateDocumentInput(docType, docOptions);
        await generateToFile(outputPath, {
          type: docType,
          options: validated,
        });
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
    const format = parseFormat(args.format as string | undefined);
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
    const format = parseFormat(positional[0]);
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

const mainCommand = defineCommand({
  meta: {
    name: "office-open",
    version: "0.10.15",
    description: "Generate Office files (.docx, .pptx, .xlsx) from JSON",
  },
  subCommands: {
    docx: createConvertCommand("docx", "docx"),
    pptx: createConvertCommand("pptx", "pptx"),
    xlsx: createConvertCommand("xlsx", "xlsx"),
    schema: schemaCommand,
  },
  args: {
    type: {
      type: "enum",
      description: "File type to generate",
      options: ["docx", "pptx", "xlsx"],
    },
  },
  async run() {
    // citty shows usage when no subcommand is matched
  },
});

void runMain(mainCommand);
