import type { EffectsOptions } from "@shared/drawingml/effects";
import type { FillOptions } from "@shared/drawingml/fill";

export interface BackgroundOptions {
  fill?: FillOptions;
  effects?: EffectsOptions;
  shadeToTitle?: boolean;
  blackWhiteMode?:
    | "clr"
    | "gray"
    | "ltGray"
    | "invGray"
    | "gmGray"
    | "bw"
    | "auto"
    | "black"
    | "white";
}
