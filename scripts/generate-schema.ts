/**
 * Generate JSON Schema (draft-07) for the public Options types of each format.
 *
 * The TS interfaces are the single source of truth for the public API; this
 * script freezes them into per-format schema files that AI tools and ajv can
 * consume directly. Output is committed at packages/office-open/schemas/.
 *
 * Post-processing passes on the raw ts-json-schema-generator output:
 *   P0  assert every $ref resolves inside definitions (catches same-name
 *       type collisions that silently overwrite definitions)
 *   P1  template-literal types (UniversalMeasure family) degrade to plain
 *       string — restore unit-constrained patterns by definition name
 *   P2  single-key tagged-union variants get a `title` naming their tag,
 *       making long anyOf lists scannable for LLMs
 *   P3  envelope: $schema/$id/title/description on the root
 *
 * Usage:
 *   pnpm schema:generate    # write schema files, then vp fmt them (root script)
 *   pnpm schema:check       # re-generate in memory and diff content, not bytes
 *   --json                  # print metrics only
 */
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

import { createGenerator } from "ts-json-schema-generator";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT_DIR = path.resolve(__dirname, "..");
const OUT_DIR = path.resolve(ROOT_DIR, "packages/office-open/schemas");

interface FormatConfig {
  /** Format label (docx | pptx | xlsx) */
  format: string;
  /** Root entry type exported by the package barrel */
  entryType: string;
  /** Package source dir (entry file = src/index.ts) */
  srcDir: string;
  title: string;
  description: string;
}

const FORMATS: FormatConfig[] = [
  {
    format: "docx",
    entryType: "DocumentOptions",
    srcDir: "packages/docx",
    title: "office-open DOCX document options",
    description:
      "Options for generating a WordprocessingML (.docx) document. " +
      "Mirrors the DocumentOptions TypeScript interface of @office-open/docx.",
  },
  {
    format: "pptx",
    entryType: "PresentationOptions",
    srcDir: "packages/pptx",
    title: "office-open PPTX presentation options",
    description:
      "Options for generating a PresentationML (.pptx) file. " +
      "Mirrors the PresentationOptions TypeScript interface of @office-open/pptx.",
  },
  {
    format: "xlsx",
    entryType: "WorkbookOptions",
    srcDir: "packages/xlsx",
    title: "office-open XLSX workbook options",
    description:
      "Options for generating a SpreadsheetML (.xlsx) workbook. " +
      "Mirrors the WorkbookOptions TypeScript interface of @office-open/xlsx.",
  },
];

/**
 * P1 overrides: ts-json-schema-generator degrades template-literal types with
 * `${number}` holes to plain `{"type":"string"}` (literal enumeration aborts on
 * the infinite number space). Restore the unit semantics as regex patterns.
 */
const PATTERN_OVERRIDES: Record<string, { pattern: string; description: string }> = {
  UniversalMeasure: {
    pattern: "^-?(\\d+(\\.\\d+)?|\\.\\d+)(mm|cm|in|pt|pc|pi|px)$",
    description: "Measurement string: number followed by a unit (mm, cm, in, pt, pc, pi or px).",
  },
  PositiveUniversalMeasure: {
    pattern: "^(\\d+(\\.\\d+)?|\\.\\d+)(mm|cm|in|pt|pc|pi)$",
    description: "Positive measurement string: number followed by mm, cm, in, pt, pc or pi.",
  },
  Percentage: {
    pattern: "^-?(\\d+(\\.\\d+)?|\\.\\d+)%$",
    description: 'Percentage string: number followed by % (e.g. "50%", "-10.5%").',
  },
  PositivePercentage: {
    pattern: "^(\\d+(\\.\\d+)?|\\.\\d+)%$",
    description: "Positive percentage string: number followed by %.",
  },
  RelativeMeasure: {
    pattern: "^-?(\\d+(\\.\\d+)?|\\.\\d+)(em|ex)$",
    description: "Relative measurement string: number followed by em or ex.",
  },
};

/**
 * P4 definition overrides: TS types whose members are runtime objects with no
 * JSON shape (typed arrays, blobs, streams). The JSON-representable inputs a
 * caller can actually write are a string or a byte array, so replace the whole
 * definition with those legs. contentEncoding:"base64" is deliberately not
 * used — the string leg also admits data: URLs and plain UTF-8 text, so the
 * encoding is stated in prose instead of a keyword that would be half-true.
 */
const DEFINITION_OVERRIDES: Record<string, Record<string, unknown>> = {
  DataType: {
    description: "Binary content as a string or an array of byte values.",
    anyOf: [
      {
        type: "string",
        description: "Base64 string, data:[<mediatype>][;base64],<data> URL, or plain text.",
      },
      { type: "array", items: { type: "integer", minimum: 0, maximum: 255 } },
    ],
  },
};

// ── Post-processing ──

/** P0: every $ref in the schema must resolve to a definition. */
function assertRefsResolve(schema: unknown, definitions: Set<string>, at = "$") {
  if (Array.isArray(schema)) {
    for (const item of schema) assertRefsResolve(item, definitions, at);
    return;
  }
  if (schema && typeof schema === "object") {
    const obj = schema as Record<string, unknown>;
    const ref = obj.$ref;
    if (typeof ref === "string" && ref.startsWith("#/definitions/")) {
      // Generic types produce definition keys like "Foo<Bar>"; the $ref is a
      // JSON-pointer fragment, so the angle brackets arrive percent-encoded.
      const name = decodeURIComponent(ref.slice("#/definitions/".length));
      if (!definitions.has(name)) {
        throw new Error(`Dangling $ref "${ref}" at ${at} — definition was likely overwritten`);
      }
    }
    for (const [key, value] of Object.entries(obj)) {
      assertRefsResolve(value, definitions, `${at}.${key}`);
    }
  }
}

/**
 * P1+P2 in one walk: apply pattern overrides to matching definitions and
 * decorate single-key anyOf variants with a title naming their tag.
 * Returns metrics for the AI-friendliness report.
 */
function postProcess(schema: Record<string, unknown>) {
  let describedProperties = 0;
  let totalProperties = 0;
  let maxAnyOfBranches = 0;

  const visit = (node: unknown) => {
    if (Array.isArray(node)) {
      for (const item of node) visit(item);
      return;
    }
    if (!node || typeof node !== "object") return;
    const obj = node as Record<string, unknown>;

    // P2a: tsj's extended JSDoc turns `(p:cNvPr @id)` in a description into a
    // spurious nested "$id" ("…). Auto-generated …"), which ajv then tries to
    // resolve as a reference host. Only the root envelope carries an $id.
    delete obj.$id;

    // P2: anyOf whose branches each require exactly one property — the
    // tagged-union shape. Title each branch after its required tag.
    if (Array.isArray(obj.anyOf)) {
      maxAnyOfBranches = Math.max(maxAnyOfBranches, obj.anyOf.length);
      for (const branch of obj.anyOf) {
        if (!branch || typeof branch !== "object") continue;
        const b = branch as Record<string, unknown>;
        const required = b.required;
        if (
          Array.isArray(required) &&
          required.length === 1 &&
          typeof required[0] === "string" &&
          b.title === undefined
        ) {
          b.title = required[0];
        }
      }
    }

    if (obj.properties && typeof obj.properties === "object") {
      for (const [name, sub] of Object.entries(obj.properties as Record<string, unknown>)) {
        totalProperties++;
        if (
          sub &&
          typeof sub === "object" &&
          typeof (sub as Record<string, unknown>).description === "string"
        ) {
          describedProperties++;
        } else if (
          sub &&
          typeof sub === "object" &&
          Array.isArray((sub as Record<string, unknown>).anyOf)
        ) {
          // union branches carry their own descriptions — count as described
          // when at least one branch has one
          const anyOf = (sub as Record<string, unknown>).anyOf as unknown[];
          if (
            anyOf.some(
              (b) =>
                b &&
                typeof b === "object" &&
                typeof (b as Record<string, unknown>).description === "string",
            )
          ) {
            describedProperties++;
          }
        } else if (name) {
          // fallthrough: undescribed property
        }
      }
    }

    for (const value of Object.values(obj)) visit(value);
  };

  visit(schema);

  // P1: pattern overrides (definition-level replacement)
  const definitions = (schema.definitions ?? {}) as Record<string, Record<string, unknown>>;
  for (const [name, replacement] of Object.entries(DEFINITION_OVERRIDES)) {
    if (definitions[name]) definitions[name] = replacement;
  }
  for (const [name, override] of Object.entries(PATTERN_OVERRIDES)) {
    const def = definitions[name];
    if (!def) continue;
    // tsj emits the degraded form as {"type":"string"} (+"description" from JSDoc)
    if (def.type === "string") {
      def.pattern = override.pattern;
      def.description = override.description;
    } else if (def.anyOf) {
      // `${number}${unit}` degrades to anyOf[number, string-of-units] in some
      // versions — replace the whole definition with the pattern form
      definitions[name] = { type: "string", pattern: override.pattern, ...override };
    }
  }

  return {
    definitions: Object.keys(definitions).length,
    totalProperties,
    describedProperties,
    descriptionCoverage:
      totalProperties > 0 ? +((describedProperties / totalProperties) * 100).toFixed(1) : 0,
    maxAnyOfBranches,
  };
}

/** Deterministic serialization: recursively sort object keys. Byte-stable, so
 * the --check mode can diff this against the committed file re-serialized the
 * same way — formatting differences never count as drift. */
function canonicalJson(value: unknown): string {
  const sort = (v: unknown): unknown => {
    if (Array.isArray(v)) return v.map(sort);
    if (v && typeof v === "object") {
      return Object.fromEntries(
        Object.entries(v as Record<string, unknown>)
          .sort(([a], [b]) => (a < b ? -1 : 1))
          .map(([k, item]) => [k, sort(item)]),
      );
    }
    return v;
  };
  return JSON.stringify(sort(value), null, 2) + "\n";
}

// ── Main ──

interface FormatResult {
  format: string;
  metrics: ReturnType<typeof postProcess>;
  output: string;
}

function generateFormat(config: FormatConfig): FormatResult {
  const generator = createGenerator({
    path: path.resolve(ROOT_DIR, config.srcDir, "src/index.ts"),
    tsconfig: path.resolve(ROOT_DIR, config.srcDir, "tsconfig.json"),
    type: config.entryType,
    // "export": named types become shared definitions referenced by $ref.
    // "none" would inline the whole closure into every use site and blow up
    // the JSON string. Expose "export" requires same-name types to be unique
    // across the closure — enforced by the source-level dedup that keeps the
    // probe (closure collision scan) at zero.
    expose: "export",
    jsDoc: "extended",
    functions: "hide",
    skipTypeCheck: false,
  });

  const schema = generator.createSchema(config.entryType) as Record<string, unknown>;

  const definitions = (schema.definitions ?? {}) as Record<string, unknown>;
  // P0: collision detection via ref resolution
  assertRefsResolve(schema, new Set(Object.keys(definitions)));

  const metrics = postProcess(schema);

  // P3: envelope — consumer-facing only; no repo-workflow wording here
  schema.$schema = "http://json-schema.org/draft-07/schema#";
  schema.$id = `https://cdn.jsdelivr.net/npm/office-open/schemas/${config.format}.schema.json`;
  schema.title = config.title;
  schema.description = config.description;

  return { format: config.format, metrics, output: canonicalJson(schema) };
}

function main() {
  const args = process.argv.slice(2);
  const checkOnly = args.includes("--check");
  const jsonOnly = args.includes("--json");

  const results: FormatResult[] = [];
  let failed = false;

  for (const config of FORMATS) {
    process.stderr.write(`  generating ${config.format}...`);
    const result = generateFormat(config);
    results.push(result);
    process.stderr.write(
      ` ${result.metrics.definitions} definitions, ` +
        `${result.metrics.descriptionCoverage}% described, ` +
        `max anyOf ${result.metrics.maxAnyOfBranches}\n`,
    );
  }

  if (jsonOnly) {
    console.log(
      JSON.stringify(
        results.map((r) => ({ format: r.format, ...r.metrics })),
        null,
        2,
      ),
    );
    return;
  }

  if (!fs.existsSync(OUT_DIR)) fs.mkdirSync(OUT_DIR, { recursive: true });

  for (const result of results) {
    const outPath = path.join(OUT_DIR, `${result.format}.schema.json`);
    if (checkOnly) {
      // Compare parsed content, not bytes: `vp fmt` may fold arrays differently
      // than the raw generator output, and that must never read as drift.
      const committed = fs.existsSync(outPath) ? fs.readFileSync(outPath, "utf-8") : null;
      const matches = committed !== null && canonicalJson(JSON.parse(committed)) === result.output;
      if (!matches) {
        console.error(`DRIFT: ${path.relative(ROOT_DIR, outPath)} does not match generated output`);
        failed = true;
      }
      continue;
    }
    fs.writeFileSync(outPath, result.output, "utf-8");
    console.log(`wrote ${path.relative(ROOT_DIR, outPath)}`);
  }

  if (checkOnly && failed) process.exit(1);
}

main();
