import type { UniversalMeasure } from "@office-open/core";
import { PresetDash } from "@office-open/core/drawingml";
import type { OutlineOptions as CoreOutlineOptions } from "@office-open/core/drawingml";

/**
 * PPTX-specific outline options (backward-compatible API).
 */
export interface OutlineOptions {
  width?: number | UniversalMeasure;
  color?: string;
  dashStyle?: "solid" | "dash" | "dashDot" | "lgDash" | "sysDot" | "sysDash";
}

export type { CoreOutlineOptions as OutlineOptionsCore };

const DASH_STYLE_MAP: Record<string, (typeof PresetDash)[keyof typeof PresetDash]> = {
  solid: "solid",
  dash: "dash",
  dashDot: "dashDot",
  lgDash: "lgDash",
  sysDot: "sysDot",
  sysDash: "sysDash",
};

/**
 * Map PPTX simplified OutlineOptions to core OutlineOptions (without arrowheads).
 */
export function toCoreOutlineOptions(options: OutlineOptions = {}): CoreOutlineOptions {
  return {
    width: options.width,
    ...(options.color
      ? { type: "solidFill" as const, color: { value: options.color.replace("#", "") } }
      : { type: "noFill" as const }),
    ...(options.dashStyle && {
      dash: DASH_STYLE_MAP[options.dashStyle] ?? "solid",
    }),
  };
}
