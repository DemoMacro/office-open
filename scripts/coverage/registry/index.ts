import { DML_CHART_REGISTRY } from "./dml-chart";
import { DML_DIAGRAM_REGISTRY } from "./dml-diagram";
import { DML_MAIN_REGISTRY } from "./dml-main";
import { DML_WP_REGISTRY } from "./dml-wp";
import { MATH_REGISTRY } from "./math";
import { PML_REGISTRY } from "./pml";
import { SML_REGISTRY } from "./sml";
import type { ModelRegistryEntry } from "./types";
import { VML_MAIN_REGISTRY } from "./vml-main";
import { WML_REGISTRY } from "./wml";

export const MODEL_REGISTRIES: readonly ModelRegistryEntry[] = [
  ...DML_CHART_REGISTRY,
  ...DML_DIAGRAM_REGISTRY,
  ...DML_MAIN_REGISTRY,
  ...DML_WP_REGISTRY,
  ...MATH_REGISTRY,
  ...PML_REGISTRY,
  ...SML_REGISTRY,
  ...WML_REGISTRY,
  ...VML_MAIN_REGISTRY,
];
