import type { OutputType } from "../opc/output";
/**
 * Shared option types for patch operations across all OOXML formats.
 *
 * Each format package (docx/pptx/xlsx) extends BasePatchOptions with its own
 * patches field and format-specific options, reusing the same data/outputType/
 * placeholderDelimiters shape so the patch API is uniform across formats.
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
 * Format-specific patch functions extend this with their `patches` field
 * (and any format-specific options like `keepOriginalStyles`).
 */
export interface BasePatchOptions<T extends OutputType = OutputType> {
  /** Source document bytes (Buffer / Uint8Array / ArrayBuffer / base64 data URL / …). */
  data: DataType;
  /** Output container type — controls the return type via OutputByType. */
  outputType: T;
  /** Custom placeholder delimiters (default `{{` / `}}`). */
  placeholderDelimiters?: PlaceholderDelimiters;
}
