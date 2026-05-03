export {
    convertPixelsToEmu as pixelsToEmus,
    convertEmuToPixels as emusToPixels,
    convertInchesToEmu as inchesToEmus,
    convertEmuToInches as emusToInches,
    convertPointsToEmu as pointsToEmus,
    convertEmuToPoints as emusToPoints,
} from "@office-open/core";

/** Convert percentage to thousandths of a percent (used in some OOXML attributes) */
export const percentToTHousandths = (percent: number): number => Math.round(percent * 1000);

/** Common slide size presets (in pixels at 96 DPI) */
export const SlideSizePreset = {
    WIDE: { width: 960, height: 540 },
    STANDARD_4X3: { width: 720, height: 540 },
    WIDESCREEN_16X10: { width: 960, height: 600 },
    WIDESCREEN_16X9: { width: 960, height: 540 },
} as const;
