/**
 * Run options type for PPTX text runs.
 *
 * @module
 */
import type { RunPropertiesOptions } from "./run-properties";

export interface RunOptions extends RunPropertiesOptions {
  text?: string;
}
