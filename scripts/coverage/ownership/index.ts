import { DML_CHART_OWNERSHIP } from "./dml-chart";
import { DML_DIAGRAM_OWNERSHIP } from "./dml-diagram";
import { DML_MAIN_OWNERSHIP } from "./dml-main";
import { DML_WP_OWNERSHIP } from "./dml-wp";
import { MATH_OWNERSHIP } from "./math";
import { PML_OWNERSHIP } from "./pml";
import { SML_OWNERSHIP } from "./sml";
import type { ModelOwnershipEntry } from "./types";
import { VML_MAIN_OWNERSHIP } from "./vml-main";
import { WML_OWNERSHIP } from "./wml";

export const MODEL_OWNERSHIP: readonly ModelOwnershipEntry[] = [
  ...DML_CHART_OWNERSHIP,
  ...DML_DIAGRAM_OWNERSHIP,
  ...DML_MAIN_OWNERSHIP,
  ...DML_WP_OWNERSHIP,
  ...MATH_OWNERSHIP,
  ...PML_OWNERSHIP,
  ...SML_OWNERSHIP,
  ...WML_OWNERSHIP,
  ...VML_MAIN_OWNERSHIP,
];
