/**
 * Per-QName OOXML ownership and runtime-evidence report.
 *
 * The existing XSD scanner proves that a QName has implementation-side
 * handling. This report additionally requires at least one spec-side reference
 * before considering that QName runtime-covered. It is the bridge between the
 * structured XSD model and the one-hot fixture matrix.
 */
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

import { extractUsedAttributes, extractUsedElements, XSD_CONFIGS } from "./xsd";
import { UNTRACKABLE_ATTRS } from "./xsd";
import { parseXsdStructured } from "./xsd-model";

const ROOT_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");

interface CoverageEntry {
  schema: string;
  qname: string;
  ownership: "implemented" | "missing" | "untrackable";
  runtime: "fixture" | "missing" | "untrackable";
}

interface AttributeEntry extends CoverageEntry {
  owner: string;
}

function main(): void {
  const elements: CoverageEntry[] = [];
  const attributes: AttributeEntry[] = [];

  for (const config of XSD_CONFIGS) {
    const schemaDirectory = path.resolve(
      ROOT_DIR,
      "ooxml-schemas",
      config.xsdDir ?? "transitional",
    );
    const schema = parseXsdStructured(config.xsdFile, schemaDirectory);
    const implementationElements = extractUsedElements(config);
    const fixtureElements = extractUsedElements(config, {
      directories: config.searchDirs,
      includeTests: true,
    });
    const implementationAttributes = extractUsedAttributes(config);
    const fixtureAttributes = extractUsedAttributes(config, {
      directories: config.searchDirs,
      includeTests: true,
    });

    for (const [localName, element] of schema.elements) {
      elements.push({
        schema: config.label,
        qname: element.qname,
        ownership: implementationElements.has(localName) ? "implemented" : "missing",
        runtime: fixtureElements.has(localName) ? "fixture" : "missing",
      });
      for (const attribute of element.attributes) {
        const localAttributeName = attribute.name.includes(":")
          ? attribute.name.split(":")[1]!
          : attribute.name;
        const untrackable = UNTRACKABLE_ATTRS.has(localAttributeName);
        attributes.push({
          schema: config.label,
          qname: `${element.qname}/@${attribute.name}`,
          owner: element.qname,
          ownership: untrackable
            ? "untrackable"
            : implementationAttributes.has(localAttributeName)
              ? "implemented"
              : "missing",
          runtime: untrackable
            ? "untrackable"
            : fixtureAttributes.has(localAttributeName)
              ? "fixture"
              : "missing",
        });
      }
    }
  }

  const elementOwnershipGaps = elements.filter((entry) => entry.ownership === "missing");
  const elementRuntimeGaps = elements.filter((entry) => entry.runtime === "missing");
  const attributeUntrackable = attributes.filter((entry) => entry.ownership === "untrackable");
  const attributeOwnershipGaps = attributes.filter((entry) => entry.ownership === "missing");
  const attributeRuntimeGaps = attributes.filter((entry) => entry.runtime === "missing");

  for (const entry of elementOwnershipGaps.slice(0, 40))
    console.log(`element ownership gap: ${entry.schema}:${entry.qname}`);
  for (const entry of elementRuntimeGaps.slice(0, 40))
    console.log(`element fixture gap: ${entry.schema}:${entry.qname}`);
  for (const entry of attributeOwnershipGaps.slice(0, 40))
    console.log(`attribute ownership gap: ${entry.schema}:${entry.qname}`);
  for (const entry of attributeRuntimeGaps.slice(0, 40))
    console.log(`attribute fixture gap: ${entry.schema}:${entry.qname}`);

  console.log(
    `element ownership: ${elements.length - elementOwnershipGaps.length}/${elements.length}; ` +
      `runtime fixtures: ${elements.length - elementRuntimeGaps.length}/${elements.length}`,
  );
  console.log(
    `attribute ownership: ${attributes.length - attributeOwnershipGaps.length}/${attributes.length}; ` +
      `runtime fixtures: ${attributes.length - attributeRuntimeGaps.length}/${attributes.length} ` +
      `(${attributeUntrackable.length} untrackable)`,
  );

  if (process.argv.includes("--json")) {
    const jsonPath = process.argv[process.argv.indexOf("--json") + 1];
    if (!jsonPath) throw new Error("--json requires a path");
    const output = {
      elements,
      attributes,
      summary: {
        elementOwnership: `${elements.length - elementOwnershipGaps.length}/${elements.length}`,
        elementRuntime: `${elements.length - elementRuntimeGaps.length}/${elements.length}`,
        attributeOwnership: `${attributes.length - attributeOwnershipGaps.length}/${attributes.length}`,
        attributeRuntime: `${attributes.length - attributeRuntimeGaps.length}/${attributes.length}`,
      },
    };
    const resolved = path.resolve(ROOT_DIR, jsonPath);
    fs.mkdirSync(path.dirname(resolved), { recursive: true });
    fs.writeFileSync(resolved, `${JSON.stringify(output, null, 2)}\n`);
  }
}

main();
