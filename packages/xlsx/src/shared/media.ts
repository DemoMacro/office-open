/**
 * Media collection for XLSX files — stores image binary data.
 *
 * The deduplicated {@link Media} collection lives in @office-open/core and is
 * re-exported here so XLSX call sites and the public API share one
 * content-based dedup implementation across all format packages.
 *
 * @module
 */

export interface MediaData {
  fileName: string;
  type: string;
  data: Uint8Array;
  width: number;
  height: number;
}

export { Media } from "@office-open/core";
