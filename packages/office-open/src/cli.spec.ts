import { execFile } from "node:child_process";
import { mkdtemp, readFile, rm, writeFile } from "node:fs/promises";
import { tmpdir } from "node:os";
import { join } from "node:path";
import { promisify } from "node:util";

import { describe, expect, it } from "vite-plus/test";

import { detectOfficeFormat } from "./formats";

const run = promisify(execFile);
const cli = join(import.meta.dirname, "../dist/cli.mjs");

interface CommandResult {
  code: number;
  stderr: string;
  stdout: string;
}

async function cliRun(...args: string[]): Promise<CommandResult> {
  try {
    const { stdout, stderr } = await run(process.execPath, [cli, ...args], {
      encoding: "utf8",
    });
    return { code: 0, stdout, stderr };
  } catch (error) {
    const failure = error as { code?: number; stdout?: string; stderr?: string };
    return {
      code: failure.code ?? 1,
      stdout: failure.stdout ?? "",
      stderr: failure.stderr ?? "",
    };
  }
}

describe("CLI", () => {
  it("shows root and schema help", async () => {
    const root = await cliRun("--help");
    expect(root.code).toBe(0);
    expect(root.stdout).toContain("detect");
    expect(root.stdout).toContain("parse");
    expect(root.stdout).toContain("schema");

    const schema = await cliRun("schema", "--help");
    expect(schema.code).toBe(0);
    expect(schema.stdout).toContain("index");
    expect(schema.stdout).toContain("slice");
  });

  it("generates, detects, and parses a DOCX file", async () => {
    const directory = await mkdtemp(join(tmpdir(), "office-open-cli-"));
    try {
      const input = join(directory, "input.json");
      const output = join(directory, "output.docx");
      const parsedOutput = join(directory, "parsed.json");
      await writeFile(
        input,
        JSON.stringify({ sections: [{ children: [{ paragraph: "CLI dispatch" }] }] }),
      );

      const generated = await cliRun("docx", input, output);
      expect(generated.code).toBe(0);
      expect(generated.stdout).toContain(`Generated: ${output}`);
      const bytes = new Uint8Array(await readFile(output));
      expect(detectOfficeFormat(bytes).format).toBe("docx");

      const parsed = await cliRun("parse", output, "-o", parsedOutput);
      expect(parsed.code).toBe(0);
      expect(JSON.parse(await readFile(parsedOutput, "utf8"))).toMatchObject({
        options: { sections: expect.any(Array) },
        type: "docx",
      });

      const detected = await cliRun("detect", output, "--json");
      expect(detected.code).toBe(0);
      expect(JSON.parse(detected.stdout)).toMatchObject({
        container: "zip",
        family: "document",
        format: "docx",
      });
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });

  it("slices schema definitions", async () => {
    const result = await cliRun("schema", "slice", "docx", "ParagraphOptions");
    expect(result.code).toBe(0);
    expect(result.stdout).toContain("ParagraphOptions");
  });

  it("returns nonzero errors for invalid commands and input", async () => {
    const directory = await mkdtemp(join(tmpdir(), "office-open-cli-error-"));
    try {
      const invalid = join(directory, "invalid.json");
      const output = join(directory, "invalid.docx");
      await writeFile(invalid, "{ sections: ");
      expect((await cliRun("docx", invalid, output)).code).toBe(1);
      expect((await cliRun("bogus-format", invalid, output)).code).toBe(1);
      expect((await cliRun("parse", join(directory, "missing.docx"))).code).toBe(1);

      const unknown = await cliRun("schema", "slice", "docx", "NotAType");
      expect(unknown.code).toBe(1);
      expect(unknown.stderr).toContain("Unknown definition");
    } finally {
      await rm(directory, { force: true, recursive: true });
    }
  });
});
