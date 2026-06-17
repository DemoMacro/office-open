/**
 * Field-consistency helpers for descriptor round-trip auditing.
 *
 * Pairs with {@link DescriptorFieldSpec}: {@link diffTagSets} flags the static
 * asymmetries between the interface / write / parse field sets (F1/F2/F3/F5),
 * and {@link roundTripFields} runs an actual stringify→parse cycle and reports
 * which fields were lost, gained, or mutated (F4).
 *
 * @module
 */
import { parse as parseXml } from "@office-open/xml";
import type { Element } from "@office-open/xml";

import type { DescriptorFieldSpec } from "./field-spec";

export interface FieldConsistencyReport {
  /** F1 — declared on the interface but never written (write-loss). */
  f1WriteLoss: readonly string[];
  /** F2 — written but absent from the interface (write-only inflation). */
  f2WriteOnly: readonly string[];
  /** F3 — written but never restored on parse (round-trip loss). */
  f3ParseLoss: readonly string[];
  /** F5 — restored on parse but never written (parse-only). */
  f5ParseOnly: readonly string[];
}

/** Diff the interface / write / parse field sets declared on a spec. */
export function diffTagSets(spec: DescriptorFieldSpec): FieldConsistencyReport {
  const interfaceSet = new Set(spec.interfaceFields);
  const writeSet = new Set(spec.writeFields);
  const parseSet = new Set(spec.parseFields);
  return {
    f1WriteLoss: spec.interfaceFields.filter((f) => !writeSet.has(f)),
    f2WriteOnly: spec.writeFields.filter((f) => !interfaceSet.has(f)),
    f3ParseLoss: spec.writeFields.filter((f) => !parseSet.has(f)),
    f5ParseOnly: spec.parseFields.filter((f) => !writeSet.has(f)),
  };
}

export interface RoundTripResult {
  /** Present in the sample, absent after round-trip. */
  lost: readonly string[];
  /** Absent from the sample, present after round-trip. */
  gained: readonly string[];
  /** Present on both sides with differing values. */
  mutated: readonly string[];
}

/**
 * Run a stringify→parse cycle over a sample and diff the result field set.
 * Returns the F4 round-trip drift. The caller supplies the descriptor's own
 * stringify/parse (a `CustomDescriptor` pair or a function pair).
 */
export function roundTripFields<T extends object, WC = unknown, RC = unknown>(
  stringifyFn: (opts: T, ctx: WC) => string | undefined,
  parseFn: (el: Element, ctx: RC) => T,
  sample: T,
  writeCtx: WC,
  readCtx: RC,
): RoundTripResult {
  const xml = stringifyFn(sample, writeCtx);
  if (xml === undefined) {
    throw new Error("roundTripFields: stringify returned undefined for the sample");
  }
  const doc = parseXml(xml) as Element;
  const root = doc.elements?.find((e) => e.type === "element");
  if (!root) throw new Error("roundTripFields: stringify produced no root element");
  const result = parseFn(root, readCtx);
  return diffObjects(sample, result);
}

export interface OrderViolation {
  /** Position of the offending child within its parent. */
  index: number;
  /** Local name of the child element (namespace prefix stripped). */
  tag: string;
  /** Why the child violates the expected sequence. */
  reason: "unexpected" | "out-of-order";
}

/**
 * Verify the child element order of a stringified container against an XSD
 * sequence — the F6 check. Namespace prefixes are normalized, so `expected`
 * entries may be qualified (`w:pStyle`) or local (`pStyle`); both match a
 * `<w:pStyle>` child. Elements absent from `expected` are flagged "unexpected";
 * a child whose expected position precedes the previous child's is
 * "out-of-order".
 */
export function checkOrder(xml: string, expected: readonly string[]): readonly OrderViolation[] {
  const doc = parseXml(xml) as Element;
  const root = doc.elements?.find((e) => e.type === "element");
  if (!root) return [];
  const expectedLocal = expected.map(localName);
  const violations: OrderViolation[] = [];
  let lastIdx = -1;
  const children = (root.elements ?? []).filter((e) => e.type === "element");
  children.forEach((child, i) => {
    const tag = localName(child.name);
    const idx = expectedLocal.indexOf(tag);
    if (idx === -1) {
      violations.push({ index: i, tag, reason: "unexpected" });
      return;
    }
    if (idx < lastIdx) {
      violations.push({ index: i, tag, reason: "out-of-order" });
    }
    lastIdx = idx;
  });
  return violations;
}

function localName(name: string | undefined): string {
  if (name === undefined) return "";
  const colon = name.lastIndexOf(":");
  return colon === -1 ? name : name.slice(colon + 1);
}

function diffObjects(sample: object, result: object): RoundTripResult {
  const sampleKeys = enumerableKeys(sample);
  const resultKeys = enumerableKeys(result);
  const lost = sampleKeys.filter((k) => !(k in result));
  const gained = resultKeys.filter((k) => !(k in sample));
  const mutated = sampleKeys.filter(
    (k) => k in result && !deepEqual(fieldValue(sample, k), fieldValue(result, k)),
  );
  return { lost, gained, mutated };
}

function enumerableKeys(value: object): string[] {
  return Object.keys(value).filter((k) => fieldValue(value, k) !== undefined);
}

function fieldValue(value: object, key: string): unknown {
  return (value as Record<string, unknown>)[key];
}

/** Structural deep equality (objects, arrays, primitives); no special-case classes. */
function deepEqual(a: unknown, b: unknown): boolean {
  if (a === b) return true;
  if (typeof a !== typeof b) return false;
  if (a === null || b === null) return a === b;
  if (typeof a !== "object") return a === b;
  const aArr = Array.isArray(a);
  const bArr = Array.isArray(b);
  if (aArr || bArr) {
    if (!aArr || !bArr) return false;
    const aa = a as unknown[];
    const bb = b as unknown[];
    if (aa.length !== bb.length) return false;
    return aa.every((v, i) => deepEqual(v, bb[i]));
  }
  const ao = a as Record<string, unknown>;
  const bo = b as Record<string, unknown>;
  const keys = new Set([...Object.keys(ao), ...Object.keys(bo)]);
  for (const k of keys) {
    if (!deepEqual(ao[k], bo[k])) return false;
  }
  return true;
}
