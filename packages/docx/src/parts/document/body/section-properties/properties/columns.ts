/**
 * Columns module for WordprocessingML section properties.
 *
 * Defines multi-column layouts within document sections.
 *
 * Reference: http://officeopenxml.com/WPsectionPr.php
 *
 * @module
 */
import type { PositiveUniversalMeasure } from "@office-open/core";

import type { ColumnAttributes } from "./column";

/**
 * Options for configuring column layout in a section.
 */
export interface ColumnsAttributes {
  space?: number | PositiveUniversalMeasure;
  count?: number;
  separate?: boolean;
  equalWidth?: boolean;
  children?: ColumnAttributes[];
}
