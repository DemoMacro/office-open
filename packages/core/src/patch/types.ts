import type { CorePropertiesOptions } from "../opc/core";
import type { OutputType } from "../opc/output";
/**
 * Shared option types for patch operations across all OOXML formats.
 *
 * Each format package (docx/pptx/xlsx) extends BasePatchOptions with its own
 * patch value type and format-specific collection edits, so the patch API is
 * uniform across formats wherever the concepts coincide.
 *
 * @module
 */
import type { DataType } from "../util/data-type";

/** Placeholder delimiter pair surrounding a patch key (e.g. `{{` / `}}`). */
export interface PlaceholderDelimiters {
  start: string;
  end: string;
}

/**
 * Shared base options for patch operations.
 *
 * `TPatch` is the format's patch value type (run-level content, block-level
 * content, or a scalar) — every occurrence of a placeholder is replaced, not
 * just the first.
 */
export interface BasePatchOptions<T extends OutputType = OutputType, TPatch = unknown> {
  /** Source document bytes (Buffer / Uint8Array / ArrayBuffer / base64 data URL / …). */
  data: DataType;
  /** Output container type — controls the return type via OutputByType. */
  outputType: T;
  /** Custom placeholder delimiters (default `{{` / `}}`). */
  placeholderDelimiters?: PlaceholderDelimiters;
  /** Placeholder substitutions: `{{key}}` (per delimiters) → patch content. */
  placeholders?: Readonly<Record<string, TPatch>>;
  /** Literal find/replace: the find string → patch content (no delimiters added). */
  findReplace?: Readonly<Record<string, TPatch>>;
  /** Core-properties metadata override (merged over the existing docProps/core.xml). */
  coreProperties?: Partial<CorePropertiesOptions>;
}
