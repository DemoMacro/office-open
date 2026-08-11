import type { EffectListOptions } from "@office-open/core/drawingml";
import type { FillOptions } from "@shared/drawingml/fill";

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
