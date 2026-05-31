import { defineCommand, runMain } from "citty";

import { generateToFile, parseInput } from "./generate";
import { validateDocumentInput } from "./schemas";

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

      const docOptions = await parseInput(jsonInput);
      const validated = validateDocumentInput(docType, docOptions);
      await generateToFile(outputPath, {
        type: docType,
        options: validated,
      });
      console.log(`Generated: ${outputPath}`);
    },
  });
}

const mainCommand = defineCommand({
  meta: {
    name: "office-open",
    version: "0.6.8",
    description: "Generate Office files (.docx, .pptx, .xlsx) from JSON",
  },
  subCommands: {
    docx: createConvertCommand("docx", "docx"),
    pptx: createConvertCommand("pptx", "pptx"),
    xlsx: createConvertCommand("xlsx", "xlsx"),
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
