export interface RawPartLike {
  path: string;
  data?: unknown;
  contentType?: string;
}

export interface RawAuditBlocker {
  part: string;
  reason: string;
}

const XML_PART = /\.(?:xml|rels|vml)$/i;

function isPlainObject(value: unknown): value is Record<string, unknown> {
  return typeof value === "object" && value !== null && Object.getPrototypeOf(value) === null;
}

function isRecordObject(value: unknown): value is Record<string, unknown> {
  return (
    typeof value === "object" &&
    value !== null &&
    !(value instanceof Uint8Array) &&
    (isPlainObject(value) || Object.getPrototypeOf(value) === Object.prototype)
  );
}

function findRawXml(value: unknown, seen: Set<object>, path: string): RawAuditBlocker[] {
  if (!isRecordObject(value)) return [];
  if (seen.has(value)) return [];
  seen.add(value);
  const blockers: RawAuditBlocker[] = [];
  for (const [key, child] of Object.entries(value)) {
    const childPath = path ? `${path}.${key}` : key;
    if (key === "rawXml" && child !== undefined) {
      blockers.push({ part: childPath, reason: "canonical-options-rawXml" });
    }
    blockers.push(...findRawXml(child, seen, childPath));
  }
  return blockers;
}

export function auditCanonicalOptions(
  options: unknown,
  allowedOpaqueParts: readonly RegExp[],
): RawAuditBlocker[] {
  const blockers = findRawXml(options, new Set(), "");
  const rawParts = (options as { rawParts?: readonly RawPartLike[] }).rawParts ?? [];
  for (const part of rawParts) {
    if (allowedOpaqueParts.some((pattern) => pattern.test(part.path))) continue;
    if (XML_PART.test(part.path)) {
      blockers.push({ part: part.path, reason: "modeled-xml-passthrough" });
      continue;
    }
    blockers.push({ part: part.path, reason: "opaque-binary-not-allowed" });
  }
  return blockers;
}
