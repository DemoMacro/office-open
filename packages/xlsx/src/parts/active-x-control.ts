/**
 * ax:ocx — ActiveX control metadata and its opaque property binary.
 *
 * @module
 */
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, escapeXml } from "@office-open/xml";
import type { Element } from "@office-open/xml";

/** ActiveX persistence mode (ax:ocx `@ax:persistence`). */
export type ActiveXPersistence =
  | "persistPropertyBag"
  | "persistStorage"
  | "persistStream"
  | "persistStreamInit";

/** Named ActiveX design-time property (ax:ocxPr). */
export interface ActiveXPropertyOptions {
  /** Property name (ax:ocxPr `@ax:name`). */
  name: string;
  /** Source property value (ax:ocxPr `@ax:value`). */
  value: string;
}

/** ActiveX control metadata (ax:ocx); binary contents remain opaque. */
export interface ActiveXControlOptions {
  /** COM class identifier, e.g. classId="{8BD21D10-EC42-11CE-9E0D-00AA006002f3}". */
  classId: string;
  /** Persistence mode, e.g. persistence="persistPropertyBag". */
  persistence?: ActiveXPersistence;
  /** License key stored verbatim, e.g. license="ABC123". */
  license?: string;
  /** Named property-bag entries (ax:ocxPr). */
  properties?: ActiveXPropertyOptions[];
  /** Opaque ActiveX property binary path (xl/activeX/activeXN.bin). */
  binaryPath?: string;
  /** Binary relationship ID (ax:ocx `@r:id`). Round-trip emission only. */
  relationshipId?: string;
}

/** Structured parse failure with the exact ActiveX location. */
export class ActiveXControlParseError extends Error {
  readonly part: string;
  readonly path: string;
  readonly name: string;
  readonly reason: string;

  constructor(part: string, path: string, name: string, reason: string) {
    super(`${part}${path}: ${name}: ${reason}`);
    this.part = part;
    this.path = path;
    this.name = name;
    this.reason = reason;
  }
}

const ACTIVE_X_NS = "http://schemas.microsoft.com/office/2006/activeX";
const ACTIVE_X_PERSISTENCE = new Set([
  "persistPropertyBag",
  "persistStorage",
  "persistStream",
  "persistStreamInit",
]);

function readActiveXControl(el: Element): ActiveXControlOptions {
  const elementName = String(el.name ?? "");
  if (elementName !== "ax:ocx") {
    throw new ActiveXControlParseError(
      "xl/activeX",
      "/ax:ocx",
      elementName,
      "unsupported root element",
    );
  }
  const classId = attr(el, "ax:classid") ?? attr(el, "classid");
  if (!classId) {
    throw new ActiveXControlParseError(
      "xl/activeX",
      "/ax:ocx",
      "classid",
      "missing required attribute",
    );
  }
  const persistence = attr(el, "ax:persistence") ?? attr(el, "persistence");
  if (persistence !== undefined && !ACTIVE_X_PERSISTENCE.has(persistence)) {
    throw new ActiveXControlParseError(
      "xl/activeX",
      "/ax:ocx",
      persistence,
      "unsupported persistence value",
    );
  }
  const license = attr(el, "ax:license") ?? attr(el, "license");
  const relationshipId = attr(el, "r:id");
  const properties = (el.elements ?? []).flatMap((child) => {
    if (child.name !== "ax:ocxPr") {
      throw new ActiveXControlParseError(
        "xl/activeX",
        "/ax:ocx",
        String(child.name ?? ""),
        "unsupported element",
      );
    }
    const name = attr(child, "ax:name") ?? attr(child, "name");
    const value = attr(child, "ax:value") ?? attr(child, "value");
    if (name === undefined || value === undefined) {
      throw new ActiveXControlParseError(
        "xl/activeX",
        "/ax:ocx/ax:ocxPr",
        String(child.name ?? ""),
        "missing required name or value",
      );
    }
    return [{ name, value }];
  });
  return {
    classId,
    ...(persistence ? { persistence: persistence as ActiveXPersistence } : {}),
    ...(license !== undefined ? { license } : {}),
    ...(relationshipId !== undefined ? { relationshipId } : {}),
    ...(properties.length > 0 ? { properties } : {}),
  };
}

function writeActiveXControl(options: ActiveXControlOptions): string {
  const attributes = [`ax:classid="${escapeXml(options.classId)}"`];
  if (options.persistence) attributes.push(`ax:persistence="${options.persistence}"`);
  if (options.license !== undefined) attributes.push(`ax:license="${escapeXml(options.license)}"`);
  if (options.relationshipId !== undefined) {
    attributes.push(`r:id="${escapeXml(options.relationshipId)}"`);
  }
  if (!options.properties?.length)
    return `<ax:ocx xmlns:ax="${ACTIVE_X_NS}" ${attributes.join(" ")}/>`;
  const children = options.properties
    .map(
      (property) =>
        `<ax:ocxPr ax:name="${escapeXml(property.name)}" ax:value="${escapeXml(property.value)}"/>`,
    )
    .join("");
  return `<ax:ocx xmlns:ax="${ACTIVE_X_NS}" ${attributes.join(" ")}>${children}</ax:ocx>`;
}

export const activeXControlDesc: CustomDescriptor<ActiveXControlOptions> = {
  kind: "custom",
  stringify(options) {
    return writeActiveXControl(options);
  },
  parse(el) {
    return readActiveXControl(el);
  },
};
