// ── Element types ──

export interface Attributes {
  [key: string]: string | number | undefined;
}

export interface DeclarationAttributes {
  version?: string | number;
  encoding?: string;
  standalone?: string;
}

export interface Element {
  declaration?: {
    attributes?: DeclarationAttributes;
  };
  instruction?: string;
  /** Body of a processing instruction without the surrounding `<?…?>`. */
  attributes?: Attributes;
  /** Verbatim values for attributes coerced by `nativeTypeAttributes`. */
  attributeRawValues?: Record<string, string>;
  cdata?: string;
  doctype?: string;
  comment?: string;
  text?: string | number | boolean;
  type?: string;
  name?: string;
  elements?: Element[];
  /**
   * Deferred content: the element's inner XML captured verbatim, children
   * not parsed (see `ParseOptions.deferElements`). Consumers that understand
   * the element scan `raw` themselves; `stringify` re-emits it verbatim, so
   * set/save round-trips keep the bytes.
   */
  raw?: string;
  parent?: Element;
}

export interface ElementCompact {
  [key: string]: unknown;
  _declaration?: {
    _attributes?: DeclarationAttributes;
  };
  _instruction?: {
    [key: string]: string;
  };
  _attributes?: Attributes;
  _cdata?: string;
  _doctype?: string;
  _comment?: string;
  _text?: string | number;
}

// ── Options: Ignore flags ──

export interface IgnoreOptions {
  ignoreDeclaration?: boolean;
  ignoreInstruction?: boolean;
  ignoreAttributes?: boolean;
  ignoreComment?: boolean;
  ignoreCdata?: boolean;
  ignoreDoctype?: boolean;
  ignoreText?: boolean;
}

// ── Options: parse ──

export interface ParseOptions extends IgnoreOptions {
  trim?: boolean;
  nativeTypeAttributes?: boolean;
  captureSpacesBetweenElements?: boolean;
  /**
   * Element names whose inner XML is captured verbatim into {@link Element.raw}
   * instead of being parsed into child elements. Lets a consumer with a
   * dedicated scanner for a hot container (e.g. xlsx `sheetData`) skip
   * materializing millions of child nodes. Same-name nesting is tracked, so
   * the capture always ends at the truly matching close tag.
   */
  deferElements?: string[];
  /**
   * Namespace URI → canonical prefix (empty string = unprefixed form).
   * Present, every prefixed element and attribute name resolves through the
   * in-scope xmlns bindings and rewrites to the canonical prefix when the
   * bound URI is known — a source that binds the Word namespace to `ns0:`
   * parses into `w:` names, so prefix-agnostic downstream matching works.
   * The xmlns declarations themselves follow the rewrite (xmlns:ns0 → xmlns:w
   * for the same URI), keeping serialized output self-consistent; unknown
   * URIs keep their source prefix verbatim. The reserved `xml` prefix and
   * unprefixed attributes (no namespace) are never rewritten.
   */
  normalizeNamespaces?: Readonly<Record<string, string>>;
}

// ── Options: stringify ──

export interface StringifyOptions extends IgnoreOptions {
  spaces?: number | string;
  indentText?: boolean;
  indentCdata?: boolean;
  fullTagEmptyElement?: boolean;
  attributeValueFn?: (
    attributeValue: string,
    attributeName: string,
    currentElementName: string,
    currentElementObj: object,
  ) => string;
}
