import type { EffectsOptions } from "@file/drawingml/effects";
import type { FillOptions } from "@file/drawingml/fill";

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
