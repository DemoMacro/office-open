import {
    createOutline,
    PresetDash,
    LineEndType,
    LineEndWidth,
    LineEndLength,
} from "@office-open/core/drawingml";
import type {
    OutlineOptions as CoreOutlineOptions,
    LineEndOptions as CoreLineEndOptions,
} from "@office-open/core/drawingml";

/**
 * PPTX-specific outline options (backward-compatible API).
 */
export interface OutlineOptions {
    readonly width?: number;
    readonly color?: string;
    readonly dashStyle?: "solid" | "dash" | "dashDot" | "lgDash" | "sysDot" | "sysDash";
}

export type { CoreOutlineOptions as OutlineOptionsCore };

const DASH_STYLE_MAP: Record<string, keyof typeof PresetDash> = {
    solid: "SOLID",
    dash: "DASH",
    dashDot: "DASH_DOT",
    lgDash: "LG_DASH",
    sysDot: "SYS_DOT",
    sysDash: "SYS_DASH",
};

const ARROWHEAD_MAP: Record<string, keyof typeof LineEndType> = {
    triangle: "TRIANGLE",
    stealth: "STEALTH",
    diamond: "DIAMOND",
    oval: "OVAL",
    open: "ARROW",
};

const ARROWHEAD_SIZE_MAP: Record<string, keyof typeof LineEndWidth> = {
    sm: "SMALL",
    med: "MEDIUM",
    lg: "LARGE",
};

function toCoreLineEnd(type: string, width?: string, length?: string): CoreLineEndOptions {
    return {
        type: ARROWHEAD_MAP[type] ?? "TRIANGLE",
        ...(width ? { width: ARROWHEAD_SIZE_MAP[width] as keyof typeof LineEndWidth } : {}),
        ...(length ? { length: ARROWHEAD_SIZE_MAP[length] as keyof typeof LineEndLength } : {}),
    };
}

/**
 * Creates an outline element using pptx's simplified API.
 */
export const createOutlineCompat = (
    options: OutlineOptions = {},
    arrowheads?: {
        readonly beginType?: string;
        readonly endType?: string;
        readonly width?: string;
        readonly length?: string;
    },
) =>
    createOutline({
        width: options.width,
        ...(options.color
            ? { type: "solidFill" as const, color: { value: options.color.replace("#", "") } }
            : { type: "noFill" as const }),
        ...(options.dashStyle && {
            dash: DASH_STYLE_MAP[options.dashStyle] ?? "SOLID",
        }),
        ...(arrowheads?.endType
            ? { headEnd: toCoreLineEnd(arrowheads.endType, arrowheads.width, arrowheads.length) }
            : {}),
        ...(arrowheads?.beginType
            ? { tailEnd: toCoreLineEnd(arrowheads.beginType, arrowheads.width, arrowheads.length) }
            : {}),
    } satisfies CoreOutlineOptions);
