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
 * `examples` give a schema-driven generator concrete instances to imitate —
 * the pattern alone leaves the unit suffix easy to miss.
 */
const PATTERN_OVERRIDES: Record<
  string,
  { pattern: string; description: string; examples?: string[] }
> = {
  UniversalMeasure: {
    pattern: "^-?(\\d+(\\.\\d+)?|\\.\\d+)(mm|cm|in|pt|pc|pi|px)$",
    description: "Measurement string: number followed by a unit (mm, cm, in, pt, pc, pi or px).",
    examples: ["1.5cm", "0.75in", "12pt"],
  },
  PositiveUniversalMeasure: {
    pattern: "^(\\d+(\\.\\d+)?|\\.\\d+)(mm|cm|in|pt|pc|pi)$",
    description: "Positive measurement string: number followed by mm, cm, in, pt, pc or pi.",
    examples: ["2.5cm", "1in", "6pt"],
  },
  Percentage: {
    pattern: "^-?(\\d+(\\.\\d+)?|\\.\\d+)%$",
    description: 'Percentage string: number followed by % (e.g. "50%", "-10.5%").',
    examples: ["50%", "-10.5%"],
  },
  PositivePercentage: {
    pattern: "^(\\d+(\\.\\d+)?|\\.\\d+)%$",
    description: "Positive percentage string: number followed by %.",
    examples: ["50%", "10.5%"],
  },
  RelativeMeasure: {
    pattern: "^-?(\\d+(\\.\\d+)?|\\.\\d+)(em|ex)$",
    description: "Relative measurement string: number followed by em or ex.",
    examples: ["1.5em", "2ex"],
  },
  // Constrained string domains (core util/values.ts) — XSD facets that no
  // enumeration can express, restored as regex on the alias definitions.
  HexColor: {
    pattern: "^[0-9A-Fa-f]{6}$",
    description: "6-digit hex RGB color without '#' (e.g. FF0000).",
    examples: ["FF0000", "4472C4"],
  },
  ArgbHexColor: {
    pattern: "^[0-9A-Fa-f]{8}$",
    description: "8-digit hex ARGB color with the alpha channel first (e.g. FF4472C4).",
    examples: ["FF4472C4", "FFED7D31"],
  },
  HexColorOrAuto: {
    pattern: "^(auto|[0-9A-Fa-f]{6})$",
    description: '"auto" or a 6-digit hex RGB color without "#".',
    examples: ["auto", "FF0000"],
  },
  Guid: {
    pattern: "^\\{[0-9A-F]{8}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{4}-[0-9A-F]{12}\\}$",
    description: "GUID in braced upper-case form (ST_Guid).",
    examples: ["{1DF903F5-8FC7-4C6A-9DF9-2AB1D10E6D84}"],
  },
  Panose: {
    pattern: "^[0-9A-Fa-f]{20}$",
    description: "20-digit hex Panose-1 font classification (ST_Panose).",
    examples: ["020F0502020204030204"],
  },
  LongHexNumber: {
    pattern: "^[0-9A-Fa-f]{8}$",
    description: "8-digit hex number (ST_LongHexNumber) — rsid or font signature.",
    examples: ["00A1B2C3"],
  },
  ShortHexNumber: {
    pattern: "^[0-9A-Fa-f]{4}$",
    description: "4-digit hex number (ST_ShortHexNumber).",
    examples: ["00A1"],
  },
  UcharHexNumber: {
    pattern: "^[0-9A-Fa-f]{2}$",
    description: "2-digit hex byte (ST_UcharHexNumber) — theme tint/shade.",
    examples: ["99"],
  },
  UnsignedShortHex: {
    pattern: "^[0-9A-Fa-f]{4}$",
    description: "4-digit hex number (ST_UnsignedShortHex) — legacy password hash.",
    examples: ["83AF"],
  },
  Base64: {
    pattern: "^[A-Za-z0-9+/]*={0,2}$",
    description: "Base64-encoded bytes (xsd:base64Binary) — password hash or salt.",
    examples: ["Hh8eLiw+KTpAPT4nPj8="],
  },
  DateTime: {
    pattern: "^-?\\d{4,}-\\d{2}-\\d{2}T\\d{2}:\\d{2}:\\d{2}(\\.\\d+)?(Z|[+-]\\d{2}:\\d{2})?$",
    description: "ISO 8601 timestamp (xsd:dateTime) — Office writes the UTC Z suffix.",
    examples: ["2024-06-01T09:30:00Z"],
  },
  RichTextColor: {
    pattern: "^([0-9A-Fa-f]{6}|[0-9A-Fa-f]{8}|[0-9]+|theme:[0-9]+)$",
    description:
      'Color as the parse encoding: 6-digit RGB ("FF0000"), 8-digit ARGB (' +
      '"FFFF0000"), a legacy palette index ("10"), or "theme:N".',
    examples: ["FF0000", "FFFF0000", "10", "theme:4"],
  },
  CnfBitmask: {
    pattern: "^[01]{12}$",
    description: "Conditional-format bit string (ST_Cnf) — exactly twelve 0/1 flags.",
    examples: ["100000000001"],
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
    examples: ["data:image/png;base64,iVBORw0KGgo="],
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

/** Resolve one pointer segment: raw key first, then URI-decoded (~0/~1 last). */
function pointerChild(container: Record<string, unknown>, segment: string): unknown {
  if (segment in container) return container[segment];
  try {
    return container[decodeURIComponent(segment.replace(/~1/g, "/").replace(/~0/g, "~"))];
  } catch {
    return undefined;
  }
}

/** P0: every $ref in the schema must resolve — name refs to a definition,
 * property-sharing pointers all the way to their target node. */
function assertRefsResolve(node: unknown, definitions: Set<string>, root: unknown, at = "$") {
  if (Array.isArray(node)) {
    for (const item of node) assertRefsResolve(item, definitions, root, at);
    return;
  }
  if (node && typeof node === "object") {
    const obj = node as Record<string, unknown>;
    const ref = obj.$ref;
    if (typeof ref === "string") {
      if (!ref.startsWith("#/definitions/")) {
        // A $ref pointing anywhere else is not a definition link. The one
        // producer seen in the wild: ts-json-schema-generator turns a JSDoc
        // "@ref" in a description into "$ref": "<text after the tag>" — keep
        // bare @attr references in backticks so they stay description text.
        throw new Error(
          `Non-definition $ref "${ref}" at ${at} — JSDoc @tag leaked into the schema`,
        );
      }
      // Generic types produce definition keys like "Foo<Bar>"; the $ref is a
      // JSON-pointer fragment, so the angle brackets arrive percent-encoded.
      const name = decodeURIComponent(ref.slice("#/definitions/".length));
      if (name.includes("/")) {
        // P8 property pointer: walk it from the document root so a broken
        // anchor fails generation here, not in a consumer at runtime.
        let target: unknown = root;
        for (const segment of ref.slice(2).split("/")) {
          target =
            target && typeof target === "object"
              ? pointerChild(target as Record<string, unknown>, segment)
              : undefined;
          if (target === undefined) break;
        }
        if (target === undefined) {
          throw new Error(`Dangling $ref "${ref}" at ${at} — pointer target missing`);
        }
      } else if (!definitions.has(name)) {
        throw new Error(`Dangling $ref "${ref}" at ${at} — definition was likely overwritten`);
      }
    }
    for (const [key, value] of Object.entries(obj)) {
      assertRefsResolve(value, definitions, root, `${at}.${key}`);
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

    // P5: JSDoc often carries verbatim XSD excerpts (## XSD: CT_* ```xml …```)
    // for developers reading the source. Schema consumers need the semantics,
    // not the XSD source — strip the fenced blocks and their heading lines;
    // drop the key entirely if nothing meaningful remains.
    if (typeof obj.description === "string" && obj.description.includes("```")) {
      const stripped = obj.description
        .replace(/```[\s\S]*?```/g, "")
        .replace(/^## XSD[^\n]*$/gm, "")
        .replace(/\n{3,}/g, "\n\n")
        .trim();
      if (stripped) obj.description = stripped;
      else delete obj.description;
    }

    // P6: provenance references (ST_*/CT_* names) are developer anchors to the
    // XSD; an LLM filling the options needs the semantics, not the spec ids.
    // Parenthesised attributions go whole; bare ids go alone; dangling
    // punctuation the removal leaves behind is tidied. Nothing else changes.
    if (typeof obj.description === "string" && /(?:ST|CT)_[A-Za-z]/.test(obj.description)) {
      const stripped = obj.description
        .replace(/\s*\((?:see |the )?(?:ST|CT)_[A-Za-z][^)]*\)/g, "")
        .replace(/\b(?:ST|CT)_[A-Za-z]+\b/g, "")
        .replace(/\(\s*\)/g, "")
        .replace(/([,;])\s*\)/g, ")")
        .replace(/\s+([.,;:)])/g, "$1")
        .replace(/([\n(])\s+[—-]\s+/g, "$1 ")
        .replace(/^\s*[—-]\s+/gm, "")
        .replace(/\s{2,}/g, " ")
        .replace(/\n{3,}/g, "\n\n")
        .trim()
        .replace(/^[.,;:]\s*/, "");
      if (stripped) obj.description = stripped;
      else delete obj.description;
    }

    // P4b: bare `Uint8Array` fields (e.g. BaseMediaEntry["data"]) are inlined
    // by tsj as the typed-array runtime shape (BYTES_PER_ELEMENT/buffer/…).
    // The JSON-representable inputs for such a field are what DataType
    // describes, so replace the whole node with a DataType reference.
    // BYTES_PER_ELEMENT alongside buffer is the typed-array fingerprint.
    const props = obj.properties as Record<string, unknown> | undefined;
    if (props && typeof props === "object" && "BYTES_PER_ELEMENT" in props && "buffer" in props) {
      for (const key of Object.keys(obj)) delete obj[key];
      Object.assign(obj, { $ref: "#/definitions/DataType" });
      return;
    }

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
        if (sub && typeof sub === "object") {
          const subRec = sub as Record<string, unknown>;
          // P7: a description that merely restates the property name carries
          // no information — drop it rather than spend tokens on an echo.
          if (typeof subRec.description === "string") {
            const norm = (t: string) => t.toLowerCase().replace(/[^a-z0-9]/g, "");
            if (norm(subRec.description) === norm(name)) delete subRec.description;
          }
        }
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
      if (override.examples) def.examples = override.examples;
    } else if (def.anyOf) {
      // `${number}${unit}` degrades to anyOf[number, string-of-units] in some
      // versions — replace the whole definition with the pattern form
      definitions[name] = { type: "string", pattern: override.pattern, ...override };
    }
  }

  const shared = shareDuplicatedProperties(definitions);

  return {
    definitions: Object.keys(definitions).length,
    totalProperties,
    describedProperties,
    descriptionCoverage:
      totalProperties > 0 ? +((describedProperties / totalProperties) * 100).toFixed(1) : 0,
    maxAnyOfBranches,
    ...shared,
  };
}

/**
 * P8: share identically-shaped property schemas beyond their first occurrence
 * via `#/definitions/<Def>/properties/<name>` pointers.
 *
 * tsj's extends handling deep-merges base-class properties into every subclass
 * definition, so shared groups (cNvPr name/id fields, revision dates, …) are
 * duplicated across dozens of definitions — measured at ~30% of the docx
 * schema. JSON Schema permits $ref to any subschema and ajv resolves internal
 * pointers, so each duplicate collapses to a pointer at the first occurrence.
 * draft-07 ignores $ref siblings, which is what sharing needs: the description
 * lives once, at the anchor, instead of echoing per copy.
 */
const SHARED_PROPERTY_MIN_BYTES = 100;

function shareDuplicatedProperties(definitions: Record<string, Record<string, unknown>>): {
  sharedProperties: number;
  sharedPropertyBytes: number;
} {
  const anchors = new Map<string, string>();
  let sharedProperties = 0;
  let sharedPropertyBytes = 0;
  const escapeSegment = (segment: string) => encodeURIComponent(segment);

  // Top-down per property: a replaced parent detaches its nested properties,
  // so the walk stops there — nested anchors only ever live under first
  // occurrences, which are never replaced.
  const walkProperties = (node: Record<string, unknown>, path: string) => {
    const properties = node.properties;
    if (properties && typeof properties === "object") {
      for (const [name, schema] of Object.entries(properties as Record<string, unknown>)) {
        if (!schema || typeof schema !== "object") continue;
        const child = schema as Record<string, unknown>;
        if (typeof child.$ref === "string") continue; // already shared
        const childPath = `${path}/properties/${escapeSegment(name)}`;
        const canonical = JSON.stringify(sortKeys(child));
        if (canonical.length >= SHARED_PROPERTY_MIN_BYTES) {
          const anchor = anchors.get(canonical);
          if (anchor !== undefined) {
            sharedProperties++;
            sharedPropertyBytes += canonical.length;
            (properties as Record<string, unknown>)[name] = { $ref: anchor };
            continue; // subtree is gone; its duplicates ride the parent's ref
          }
          anchors.set(canonical, `#${childPath}`);
        }
        walkProperties(child, childPath);
      }
    }
    // union branches and array items also carry properties containers —
    // walk them as containers (not as shareable properties themselves)
    const branches = (node.anyOf ?? node.oneOf) as unknown[] | undefined;
    if (Array.isArray(branches)) {
      branches.forEach((branch, i) => {
        if (branch && typeof branch === "object")
          walkProperties(branch as Record<string, unknown>, `${path}/anyOf/${i}`);
      });
    }
    const items = node.items;
    if (items && typeof items === "object")
      walkProperties(items as Record<string, unknown>, `${path}/items`);
  };

  for (const [name, def] of Object.entries(definitions)) {
    walkProperties(def, `/definitions/${escapeSegment(name)}`);
  }
  return { sharedProperties, sharedPropertyBytes };
}

/** Recursively sort object keys; shared by canonical serialization and the
 * canonical keys of P8 property sharing (key order must never decide whether
 * two schemas are "identical"). */
function sortKeys(value: unknown): unknown {
  if (Array.isArray(value)) return value.map(sortKeys);
  if (value && typeof value === "object") {
    return Object.fromEntries(
      Object.entries(value as Record<string, unknown>)
        .sort(([a], [b]) => (a < b ? -1 : 1))
        .map(([k, item]) => [k, sortKeys(item)]),
    );
  }
  return value;
}

/** Deterministic serialization: recursively sort object keys. Byte-stable, so
 * the --check mode can diff this against the committed file re-serialized the
 * same way — formatting differences never count as drift. */
function canonicalJson(value: unknown): string {
  return JSON.stringify(sortKeys(value), null, 2) + "\n";
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
  assertRefsResolve(schema, new Set(Object.keys(definitions)), schema);

  const metrics = postProcess(schema);
  // P0 again after post-processing: P4b rewrites nodes into DataType refs and
  // P8 introduces property pointers — both must resolve like generator refs.
  assertRefsResolve(schema, new Set(Object.keys(definitions as Record<string, unknown>)), schema);

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
        `max anyOf ${result.metrics.maxAnyOfBranches}, ` +
        `${result.metrics.sharedProperties} shared properties ` +
        `(−${Math.round(result.metrics.sharedPropertyBytes / 1024)} KB)\n`,
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
