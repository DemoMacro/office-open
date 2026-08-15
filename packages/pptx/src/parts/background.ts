import type { EffectListOptions } from "@office-open/core/drawing";
import type { FillOptions } from "@shared/drawing/fill";

export interface BackgroundOptions {
  fill?: FillOptions;
  effects?: EffectListOptions;
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
