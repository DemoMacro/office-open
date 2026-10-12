import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

import { strFromU8, unzipSync } from "fflate";
import { XmlDocument, XsdValidator } from "libxml2-wasm";
import { xmlRegisterInputProvider } from "libxml2-wasm";

import { generateDocumentSync, parseDocumentSync } from "../../packages/docx/dist/index.mjs";
import {
  generatePresentationSync,
  parsePresentationSync,
} from "../../packages/pptx/dist/index.mjs";
import { generateWorkbookSync, parseWorkbookSync } from "../../packages/xlsx/dist/index.mjs";
import { parse as parseXml } from "../../packages/xml/dist/index.mjs";
import { OOXML_EXECUTABLE_REGISTRY, type OoxmlExecutableCaseInput } from "./ooxml-fixtures";
import { XSD_CONFIGS } from "./xsd";
import { parseXsdStructured } from "./xsd-model";

const ROOT_DIR = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const EVIDENCE_PATH = path.join(ROOT_DIR, ".temp", "ooxml-executable-evidence.json");

export interface OoxmlExecutableFailure {
  caseId: string;
  reason: string;
}

export interface OoxmlExecutableReport {
  mode: "executable";
  elementOwnership: string;
  attributeOwnership: string;
  executedCases: number;
  roundTripFailures: OoxmlExecutableFailure[];
  semanticFailures: OoxmlExecutableFailure[];
  unknownFixtures: string[];
  missingElementOwnership: string[];
  missingAttributeOwnership: string[];
  evidence: OoxmlExecutableEvidenceManifest;
}

export interface OoxmlExecutableEvidenceRecord {
  caseId: string;
  schema: string;
  family: OoxmlExecutableCaseInput["family"];
  element: string;
  attributes: readonly string[];
  coveredElements: readonly string[];
  coveredAttributes: readonly string[];
  status: "passed" | "failed";
  reason?: string;
}

export interface OoxmlExecutableEvidenceManifest {
  schemaVersion: 1;
  records: readonly OoxmlExecutableEvidenceRecord[];
}

function assert(condition: unknown, message: string): asserts condition {
  if (!condition) throw new Error(message);
}

function packageXml(data: Uint8Array, partPath: string): string {
  const files = unzipSync(data);
  const part = files[partPath];
  assert(part, `${partPath} was not generated`);
  return strFromU8(part);
}

function equal(actual: unknown, expected: unknown): boolean {
  return stableStringify(actual) === stableStringify(expected);
}

function stableStringify(value: unknown): string {
  if (Array.isArray(value)) return `[${value.map((entry) => stableStringify(entry)).join(",")}]`;
  if (value === null || typeof value !== "object") return JSON.stringify(value) ?? "null";
  const entries = Object.entries(value)
    .filter(([, entry]) => entry !== undefined)
    .sort(([left], [right]) => left.localeCompare(right))
    .map(([key, entry]) => `${JSON.stringify(key)}:${stableStringify(entry)}`);
  return `{${entries.join(",")}}`;
}

function schemaAttributeKeys(attribute: string): readonly string[] {
  const separator = attribute.lastIndexOf("/@");
  if (separator < 0) return [attribute];
  const element = attribute.slice(0, separator);
  const name = attribute.slice(separator + 2);
  const localName = name.includes(":") ? name.split(":").at(-1)! : name;
  return [attribute, `${element}/@${localName}`];
}

const MC_NAMESPACE = "http://schemas.openxmlformats.org/markup-compatibility/2006";
const WORD_2013_NAMESPACE = "http://schemas.microsoft.com/office/word/2012/wordml";
const WORD_2010_NAMESPACE = "http://schemas.microsoft.com/office/word/2010/wordml";

function extractIgnorableNamespaces(xml: string): Set<string> {
  const namespaces = new Set<string>([WORD_2013_NAMESPACE, WORD_2010_NAMESPACE]);
  const ignorable = xml.match(/mc:Ignorable="([^"]+)"/)?.[1];
  if (!ignorable) return namespaces;
  for (const prefix of ignorable.split(/\s+/)) {
    const namespace = xml.match(new RegExp(`xmlns:${prefix}="([^"]+)"`))?.[1];
    if (namespace) namespaces.add(namespace);
  }
  return namespaces;
}

function isIgnorableError(message: string, ignorableNamespaces: Set<string>): boolean {
  if (message.includes(`{${MC_NAMESPACE}}Ignorable`)) return true;
  if (message.includes(`{${MC_NAMESPACE}}`) && message.includes("not expected")) return true;
  for (const namespace of ignorableNamespaces) {
    if (message.includes(`{${namespace}}`) || message.includes(`'${namespace}'`)) return true;
  }
  return false;
}

function schemaInputPath(filename: string): string | undefined {
  try {
    const url = new URL(filename);
    if (url.protocol === "file:") {
      const pathname =
        url.pathname.startsWith("/") && process.platform === "win32"
          ? url.pathname.slice(1)
          : url.pathname;
      return fs.existsSync(pathname) ? pathname : undefined;
    }
  } catch {
    // Relative XSD imports arrive as plain paths.
  }
  const candidates = [
    path.resolve(filename),
    path.resolve(ROOT_DIR, "ooxml-schemas/transitional", path.basename(filename)),
    path.resolve(ROOT_DIR, "ooxml-schemas/microsoft", path.basename(filename)),
  ];
  return candidates.find((candidate) => fs.existsSync(candidate));
}

function registerSchemaInputProvider(): void {
  xmlRegisterInputProvider({
    match(filename) {
      return schemaInputPath(filename) !== undefined;
    },
    open(filename) {
      const resolved = schemaInputPath(filename);
      if (!resolved) return undefined;
      try {
        return fs.openSync(resolved, "r");
      } catch {
        return undefined;
      }
    },
    read(fd, buffer) {
      try {
        return fs.readSync(fd, buffer, 0, buffer.byteLength, null);
      } catch {
        return -1;
      }
    },
    close(fd) {
      try {
        fs.closeSync(fd);
      } catch {
        // libxml2 treats close as best-effort cleanup.
      }
      return true;
    },
  });
}

function validateOoxmlPart(xml: string, schemaFile: string, schemaDirectory: string): void {
  registerSchemaInputProvider();
  const absoluteDirectory = path.resolve(ROOT_DIR, schemaDirectory);
  const isExtensionSchema = path.basename(absoluteDirectory) === "microsoft";
  let temporarySchema: string | undefined;
  if (isExtensionSchema) {
    const temporaryDirectory = path.join(ROOT_DIR, ".temp", "coverage", "xsd");
    fs.mkdirSync(temporaryDirectory, { recursive: true });
    temporarySchema = path.join(temporaryDirectory, path.basename(schemaFile));
    const extensionSchema = fs.readFileSync(path.join(absoluteDirectory, schemaFile), "utf-8");
    const extensionNamespace = extensionSchema.match(/targetNamespace="([^"]+)"/)?.[1];
    if (!extensionNamespace) throw new Error(`${schemaFile}: no targetNamespace`);
    const transitionalLocation = path
      .relative(temporaryDirectory, path.join(ROOT_DIR, "ooxml-schemas", "transitional"))
      .replaceAll("\\", "/");
    const extensionLocation = path
      .relative(temporaryDirectory, path.join(absoluteDirectory, schemaFile))
      .replaceAll("\\", "/");
    fs.writeFileSync(
      temporarySchema,
      `<xsd:schema xmlns:xsd="http://www.w3.org/2001/XMLSchema">` +
        `<xsd:import namespace="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ` +
        `schemaLocation="${transitionalLocation}/wml.xsd"/>` +
        `<xsd:import namespace="${extensionNamespace}" ` +
        `schemaLocation="${extensionLocation}"/></xsd:schema>`,
    );
    schemaDirectory = path.relative(ROOT_DIR, temporaryDirectory);
  }
  const previousWorkingDirectory = process.cwd();
  process.chdir(path.resolve(ROOT_DIR, schemaDirectory));
  const schemaDocument = XmlDocument.fromBuffer(
    fs.readFileSync(path.join(ROOT_DIR, schemaDirectory, schemaFile)),
  );
  const validator = XsdValidator.fromDoc(schemaDocument);
  schemaDocument.dispose();
  process.chdir(previousWorkingDirectory);

  const document = XmlDocument.fromString(xml);
  try {
    validator.validate(document);
  } catch (error) {
    const ignorableNamespaces = extractIgnorableNamespaces(xml);
    const details = (error as { details?: Array<{ message: string }> }).details?.filter(
      (detail) => !isIgnorableError(detail.message, ignorableNamespaces),
    );
    if (details?.length === 0) return;
    const message =
      details?.map((detail) => detail.message.trim()).join("\n") ?? (error as Error).message;
    throw new Error(`XSD validation failed: ${message}`);
  } finally {
    document.dispose();
    validator.dispose();
    if (temporarySchema) fs.rmSync(temporarySchema, { force: true });
  }
}

export function executeOoxmlCase(fixtureCase: OoxmlExecutableCaseInput): {
  xml: string;
  reparsedXml: string;
} {
  if (fixtureCase.family === "descriptor") {
    const body = fixtureCase.descriptor.stringify(fixtureCase.options);
    assert(typeof body === "string" && body.length > 0, `${fixtureCase.id} did not stringify`);
    const applyNamespaceAttributes = (input: string): string =>
      fixtureCase.namespaceAttributes
        ? input.replace(/^<([A-Za-z0-9_.:-]+)/, `<$1 ${fixtureCase.namespaceAttributes}`)
        : input;
    const xml = applyNamespaceAttributes(body);
    const document = parseXml(xml);
    const root = document.elements?.[0];
    assert(root, `${fixtureCase.id} did not parse to a root element`);
    const reparsed = fixtureCase.descriptor.parse(root);
    if (!equal(reparsed, fixtureCase.expected)) {
      throw new Error(
        `${fixtureCase.id} projection changed: ${stableStringify(reparsed)} !== ` +
          stableStringify(fixtureCase.expected),
      );
    }
    const reparsedBody = fixtureCase.descriptor.stringify(reparsed as never);
    assert(typeof reparsedBody === "string", `${fixtureCase.id} did not restringify`);
    const reparsedXml = applyNamespaceAttributes(reparsedBody);
    assert(reparsedXml === xml, `${fixtureCase.id} is not deterministic after parse`);
    return { xml, reparsedXml };
  }
  const output =
    fixtureCase.family === "document"
      ? generateDocumentSync(fixtureCase.options, { type: "uint8array" })
      : fixtureCase.family === "workbook"
        ? generateWorkbookSync(fixtureCase.options, { type: "uint8array" })
        : generatePresentationSync(fixtureCase.options, { type: "uint8array" });
  const xml = packageXml(output, fixtureCase.partPath);
  const validationXml = fixtureCase.validationTransform?.(xml) ?? xml;
  const elementNames = new Set([
    fixtureCase.element,
    ...fixtureCase.attributes.map((attribute) => attribute.split("/")[0]!),
  ]);
  for (const qname of elementNames) {
    assert(validationXml.includes(`<${qname}`), `${fixtureCase.id} did not generate <${qname}>`);
  }

  const reparsed =
    fixtureCase.family === "document"
      ? parseDocumentSync(output)
      : fixtureCase.family === "workbook"
        ? parseWorkbookSync(output)
        : parsePresentationSync(output);
  const actual = fixtureCase.projection(reparsed);
  if (!equal(actual, fixtureCase.expected)) {
    throw new Error(
      `${fixtureCase.id} projection changed: ${stableStringify(actual)} !== ` +
        stableStringify(fixtureCase.expected),
    );
  }

  const reparsedOutput =
    fixtureCase.family === "document"
      ? generateDocumentSync(reparsed, { type: "uint8array" })
      : fixtureCase.family === "workbook"
        ? generateWorkbookSync(reparsed, { type: "uint8array" })
        : generatePresentationSync(reparsed, { type: "uint8array" });
  const reparsedXml = packageXml(reparsedOutput, fixtureCase.partPath);
  if (xml !== reparsedXml) {
    throw new Error(`${fixtureCase.id} is not deterministic after parse`);
  }
  return { xml, reparsedXml };
}

export function buildOoxmlExecutableReport(schemaLabels: readonly string[]): OoxmlExecutableReport {
  const knownLabels = new Set(XSD_CONFIGS.map((config) => config.label));
  const configs = XSD_CONFIGS.filter((config) => schemaLabels.includes(config.label));
  const unknownFixtures = schemaLabels.filter((label) => !knownLabels.has(label));
  const elementOwnership = new Map<string, Set<string>>();
  const attributeOwnership = new Map<string, Set<string>>();
  const schemaElements = new Map<string, Set<string>>();
  const schemaAttributes = new Map<string, Set<string>>();

  for (const config of configs) {
    const schemaDirectory = `ooxml-schemas/${config.xsdDir ?? "transitional"}`;
    const schema = parseXsdStructured(
      config.xsdFile,
      schemaDirectory,
      config.prefix.replace(/:$/, ""),
    );
    const elements = new Set<string>();
    const attributes = new Set<string>();
    for (const element of schema.elements.values()) {
      elements.add(element.qname);
      for (const attribute of element.attributes) {
        attributes.add(`${element.qname}/@${attribute.name}`);
      }
    }
    schemaElements.set(config.label, elements);
    schemaAttributes.set(config.label, attributes);
    elementOwnership.set(config.label, new Set());
    attributeOwnership.set(config.label, new Set());
  }

  const roundTripFailures: OoxmlExecutableFailure[] = [];
  const semanticFailures: OoxmlExecutableFailure[] = [];
  const evidenceRecords: OoxmlExecutableEvidenceRecord[] = [];
  let executedCases = 0;

  for (const fixtureCase of OOXML_EXECUTABLE_REGISTRY.cases) {
    if (!schemaElements.has(fixtureCase.schema)) continue;
    const coveredElements = new Set<string>();
    const coveredAttributes = new Set<string>();
    try {
      const { xml: runtimeXml } = executeOoxmlCase(fixtureCase);
      const xml = fixtureCase.validationTransform?.(runtimeXml) ?? runtimeXml;
      const config = configs.find((candidate) => candidate.label === fixtureCase.schema)!;
      validateOoxmlPart(
        xml,
        fixtureCase.validationSchema ?? config.xsdFile,
        fixtureCase.validationDirectory ?? `ooxml-schemas/${config.xsdDir ?? "transitional"}`,
      );
      executedCases += 1;
      const document = parseXml(xml);
      const visit = (element: typeof document): void => {
        coveredElements.add(element.name);
        for (const [name] of Object.entries(element.attributes ?? {})) {
          coveredAttributes.add(`${element.name}/@${name}`);
        }
        for (const child of element.elements ?? []) visit(child);
      };
      for (const element of document.elements ?? []) visit(element);
      elementOwnership.get(fixtureCase.schema)?.add(fixtureCase.element);
      const ownedAttributes = attributeOwnership.get(fixtureCase.schema)!;
      for (const attribute of fixtureCase.attributes) {
        const matched = schemaAttributeKeys(attribute).find((key) =>
          schemaAttributes.get(fixtureCase.schema)?.has(key),
        );
        if (matched) {
          ownedAttributes.add(matched);
        }
      }
      for (const attribute of coveredAttributes) {
        const matched = schemaAttributeKeys(attribute).find((key) =>
          schemaAttributes.get(fixtureCase.schema)?.has(key),
        );
        if (matched) {
          ownedAttributes.add(matched);
        }
      }
      for (const match of xml.matchAll(/<([A-Za-z0-9_.:-]+)[\s/>]/g)) {
        const name = match[1]!;
        if (schemaElements.get(fixtureCase.schema)?.has(name)) {
          elementOwnership.get(fixtureCase.schema)?.add(name);
        }
      }
      evidenceRecords.push({
        caseId: fixtureCase.id,
        schema: fixtureCase.schema,
        family: fixtureCase.family,
        element: fixtureCase.element,
        attributes: fixtureCase.attributes,
        coveredElements: [...coveredElements].sort(),
        coveredAttributes: [...coveredAttributes].sort(),
        status: "passed",
      });
    } catch (error) {
      const reason = error instanceof Error ? error.message : String(error);
      const failure = { caseId: fixtureCase.id, reason };
      evidenceRecords.push({
        caseId: fixtureCase.id,
        schema: fixtureCase.schema,
        family: fixtureCase.family,
        element: fixtureCase.element,
        attributes: fixtureCase.attributes,
        coveredElements: [...coveredElements].sort(),
        coveredAttributes: [...coveredAttributes].sort(),
        status: "failed",
        reason,
      });
      if (reason.includes("projection changed")) semanticFailures.push(failure);
      else roundTripFailures.push(failure);
    }
  }

  const totalElements = [...schemaElements.values()].reduce((sum, values) => sum + values.size, 0);
  const ownedElements = [...elementOwnership.values()].reduce(
    (sum, values) => sum + values.size,
    0,
  );
  const totalAttributes = [...schemaAttributes.values()].reduce(
    (sum, values) => sum + values.size,
    0,
  );
  const ownedAttributes = [...attributeOwnership.values()].reduce(
    (sum, values) => sum + values.size,
    0,
  );
  const missingElementOwnership = [...schemaElements.entries()].flatMap(([schema, values]) =>
    [...values].filter((qname) => !elementOwnership.get(schema)?.has(qname)),
  );
  const missingAttributeOwnership = [...schemaAttributes.entries()].flatMap(([schema, values]) =>
    [...values].filter((qname) => !attributeOwnership.get(schema)?.has(qname)),
  );

  return {
    mode: "executable",
    elementOwnership: `${ownedElements}/${totalElements}`,
    attributeOwnership: `${ownedAttributes}/${totalAttributes}`,
    executedCases,
    roundTripFailures,
    semanticFailures,
    unknownFixtures,
    missingElementOwnership,
    missingAttributeOwnership,
    evidence: { schemaVersion: 1, records: evidenceRecords },
  };
}

function main(): void {
  const labels = process.argv.slice(2).length
    ? process.argv.slice(2)
    : [
        "w16cid",
        "w16cex",
        "w15",
        "w14",
        "math",
        "dml-lc",
        "dml-pic",
        "dml-cdr",
        "dml-xdr",
        "dml-wp",
        "dml-diagram",
      ];
  const report = buildOoxmlExecutableReport(labels);
  fs.mkdirSync(path.dirname(EVIDENCE_PATH), { recursive: true });
  fs.writeFileSync(EVIDENCE_PATH, `${JSON.stringify(report.evidence, null, 2)}\n`);
  for (const qname of report.unknownFixtures) console.log(`unknown fixture schema: ${qname}`);
  for (const qname of report.missingElementOwnership)
    console.log(`element execution gap: ${qname}`);
  for (const qname of report.missingAttributeOwnership)
    console.log(`attribute execution gap: ${qname}`);
  for (const failure of [...report.roundTripFailures, ...report.semanticFailures]) {
    console.log(`${failure.caseId}: ${failure.reason}`);
  }
  console.log(`coverage mode: ${report.mode}`);
  console.log(`element execution: ${report.elementOwnership}`);
  console.log(`attribute execution: ${report.attributeOwnership}`);
  console.log(`executed cases: ${report.executedCases}`);
  const failed =
    report.unknownFixtures.length > 0 ||
    report.missingElementOwnership.length > 0 ||
    report.missingAttributeOwnership.length > 0 ||
    report.roundTripFailures.length > 0 ||
    report.semanticFailures.length > 0;
  if (failed) process.exitCode = 1;
}

const invokedDirectly = process.argv[1]?.endsWith("ooxml-executable.ts");
if (invokedDirectly) main();
