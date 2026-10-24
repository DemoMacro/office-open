import * as fs from "node:fs";

import { verify, type LegacyFormat } from "./legacy-verify";

type Outcome = "pass" | "encrypted" | "invalid" | "unexpected";

interface WorkerResult {
  readonly [format in LegacyFormat]: {
    total: number;
    pass: number;
    encrypted: number;
    invalid: number;
    unexpected: number;
    reasons: Record<string, number>;
  };
}

const emptyFormatResult = () => ({
  total: 0,
  pass: 0,
  encrypted: 0,
  invalid: 0,
  unexpected: 0,
  reasons: {},
});

const emptyResult = (): WorkerResult => ({
  doc: emptyFormatResult(),
  xls: emptyFormatResult(),
  ppt: emptyFormatResult(),
  rtf: emptyFormatResult(),
  odt: emptyFormatResult(),
  ods: emptyFormatResult(),
  odp: emptyFormatResult(),
});

function classify(error: unknown): Exclude<Outcome, "pass"> {
  const diagnostic = error as {
    context?: { reason?: string };
    reason?: string;
  };
  const reason = diagnostic.context?.reason ?? diagnostic.reason;
  const message = String((error as Error)?.message ?? error);
  if (/encrypted/i.test(message)) return "encrypted";
  if (/Malformed XML: invalid character/u.test(message)) return "invalid";
  if (
    reason === "out-of-range" ||
    reason === "invalid-repeat-count" ||
    reason === "invalid-style-coverage-total" ||
    reason === "malformed-formula"
  ) {
    return "invalid";
  }
  if (
    /(?:bad signature|truncated|must begin with|invalid CFB|invalid ODF|invalid zip|invalid URL|not a supported Compound|signature is not|malformed formula|unsupported formula token 0x00|unsupported formula token 0x58|unsupported RTF destination|objdata requires binary|Manifest does not declare|Manifest declares missing package path|embedded object subdocument is missing)/i.test(
      message,
    ) ||
    /ODF package is missing or cannot read content\.xml/u.test(message)
  ) {
    return "invalid";
  }
  return "unexpected";
}

function failureCategory(error: unknown): string {
  const diagnostic = error as { context?: Record<string, unknown> };
  const details = diagnostic.context
    ? ` [${safeJson(diagnostic.context)}]`
    : diagnostic.part || diagnostic.path || diagnostic.reason
      ? ` [${safeJson({
          part: diagnostic.part,
          path: diagnostic.path,
          name: diagnostic.name,
          reason: diagnostic.reason,
        })}]`
      : "";
  return (String((error as Error)?.message ?? error) + details)
    .replace(/\s+/g, " ")
    .replace(/(["']).*?\1/g, "$1…$1")
    .replace(/\b[0-9a-f]{8,}\b/gi, "…")
    .replace(/\b\d+(?:\.\d+)?\b/g, "N")
    .trim()
    .slice(0, 160);
}

function safeJson(value: unknown): string {
  try {
    return JSON.stringify(value) ?? "{}";
  } catch {
    return '{"serialization":"unavailable"}';
  }
}

let input = "";
for await (const chunk of process.stdin) input += chunk;

const result = emptyResult();
for (const line of input.split("\n")) {
  if (!line) continue;
  const item = JSON.parse(line) as { file: string; format: LegacyFormat };
  const formatResult = result[item.format];
  formatResult.total++;
  const originalConsole = {
    log: console.log,
    info: console.info,
    warn: console.warn,
    error: console.error,
  };
  console.log = () => {};
  console.info = () => {};
  console.warn = () => {};
  console.error = () => {};
  try {
    await verify(item.format, new Uint8Array(fs.readFileSync(item.file)));
    formatResult.pass++;
  } catch (error) {
    let outcome: Exclude<Outcome, "pass">;
    try {
      outcome = classify(error);
    } catch {
      outcome = "unexpected";
    }
    formatResult[outcome]++;
    let category: string;
    try {
      category = failureCategory(error);
    } catch {
      category = "diagnostic serialization failed";
    }
    const reason = `${outcome}: ${category}`;
    formatResult.reasons[reason] = (formatResult.reasons[reason] ?? 0) + 1;
  } finally {
    console.log = originalConsole.log;
    console.info = originalConsole.info;
    console.warn = originalConsole.warn;
    console.error = originalConsole.error;
  }
}

process.stdout.write(JSON.stringify(result));
