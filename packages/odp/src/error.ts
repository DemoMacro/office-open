import type { Element } from "@office-open/xml";

/** Structured error raised when an ODP package or element cannot be parsed. */
export class OdpParseError extends Error {
  constructor(
    message: string,
    readonly part: string = "",
    readonly path: string = "",
    readonly elementName: string = "",
    readonly reason: string = "",
    options?: ErrorOptions,
  ) {
    super(message, options);
    this.name = "OdpParseError";
  }
}

export function unknownSlideChild(element: Element, parentPath: string): OdpParseError {
  const name = element.name ?? "(unknown)";
  const path = `${parentPath}/${name}`;
  const reason = "element has no canonical PresentationOptions mapping";
  return new OdpParseError(
    `content.xml: ${path}: ${name}: ${reason}`,
    "content.xml",
    path,
    name,
    reason,
  );
}

/** Structured error raised when PresentationOptions cannot be represented as ODF. */
export class OdpGenerateError extends Error {
  constructor(
    message: string,
    readonly part: string,
    readonly path: string,
    readonly elementName: string,
    readonly reason: string,
    options?: ErrorOptions,
  ) {
    super(message, options);
    this.name = "OdpGenerateError";
  }
}
