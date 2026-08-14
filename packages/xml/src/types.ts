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
  attributes?: Attributes;
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
  compact?: boolean;
  trim?: boolean;
  sanitize?: boolean;
  nativeType?: boolean;
  nativeTypeAttributes?: boolean;
  addParent?: boolean;
  alwaysArray?: boolean | string[];
  alwaysChildren?: boolean;
  instructionHasAttributes?: boolean;
  captureSpacesBetweenElements?: boolean;
  /**
   * Element names whose inner XML is captured verbatim into {@link Element.raw}
   * instead of being parsed into child elements. Lets a consumer with a
   * dedicated scanner for a hot container (e.g. xlsx `sheetData`) skip
   * materializing millions of child nodes. Same-name nesting is tracked, so
   * the capture always ends at the truly matching close tag.
   */
  deferElements?: string[];
  doctypeFn?: (value: string, parentElement: object) => string;
  instructionFn?: (value: string, instructionName: string, parentElement: string) => string;
  cdataFn?: (value: string, parentElement: object) => string;
  commentFn?: (value: string, parentElement: object) => string;
  textFn?: (value: string, parentElement: object) => string;
  instructionNameFn?: (
    instructionName: string,
    instructionValue: string,
    parentElement: string,
  ) => string;
  elementNameFn?: (value: string, parentElement: object) => string;
  attributeNameFn?: (
    attributeName: string,
    attributeValue: string,
    parentElement: string,
  ) => string;
  attributeValueFn?: (
    attributeValue: string,
    attributeName: string,
    parentElement: string,
  ) => string;
  attributesFn?: (value: Attributes, parentElement: string) => Attributes;
}

// ── Options: stringify ──

export interface StringifyOptions extends IgnoreOptions {
  spaces?: number | string;
  compact?: boolean;
  indentText?: boolean;
  indentCdata?: boolean;
  indentAttributes?: boolean;
  indentInstruction?: boolean;
  fullTagEmptyElement?: boolean;
  noQuotesForNativeAttributes?: boolean;
  doctypeFn?: (value: string, currentElementName: string, currentElementObj: object) => string;
  instructionFn?: (
    instructionValue: string,
    instructionName: string,
    currentElementName: string,
    currentElementObj: object,
  ) => string;
  cdataFn?: (value: string, currentElementName: string, currentElementObj: object) => string;
  commentFn?: (value: string, currentElementName: string, currentElementObj: object) => string;
  textFn?: (value: string, currentElementName: string, currentElementObj: object) => string;
  instructionNameFn?: (
    instructionName: string,
    instructionValue: string,
    currentElementName: string,
    currentElementObj: object,
  ) => string;
  elementNameFn?: (value: string, currentElementName: string, currentElementObj: object) => string;
  attributeNameFn?: (
    attributeName: string,
    attributeValue: string,
    currentElementName: string,
    currentElementObj: object,
  ) => string;
  attributeValueFn?: (
    attributeValue: string,
    attributeName: string,
    currentElementName: string,
    currentElementObj: object,
  ) => string;
  attributesFn?: (
    value: Attributes,
    currentElementName: string,
    currentElementObj: object,
  ) => Attributes;
  fullTagEmptyElementFn?: (currentElementName: string, currentElementObj: object) => boolean;
}
