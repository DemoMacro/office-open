import type { EffectsOptions } from "@file/drawingml/effects";
import type { FillOptions } from "@file/drawingml/fill";

export interface BackgroundOptions {
  readonly fill?: FillOptions;
  readonly effects?: EffectsOptions;
  readonly shadeToTitle?: boolean;
  readonly blackWhiteMode?:
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
