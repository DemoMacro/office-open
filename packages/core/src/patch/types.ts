/**
 * Shared option types for canonical patch commands across OOXML formats.
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
 * `TCommand` is the format-specific canonical command union.
 */
export interface PatchOptions<TCommand = unknown> {
  /** Source package in any `DataType` shape; async patch entries also accept `Blob` and `ReadableStream`. */
  data: DataType;
  /** Ordered canonical-model edits; each command sees the previous command's result. */
  commands: readonly TCommand[];
  /** Custom placeholder delimiters for `placeholder` targets (default `{{` / `}}`). */
  placeholderDelimiters?: PlaceholderDelimiters;
}
