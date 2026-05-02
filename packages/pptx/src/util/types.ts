/** EMU (English Metric Unit) conversion: 1 inch = 914400 EMUs, 1 pixel = 9525 EMUs (96 DPI) */
export const EMU_PER_PIXEL = 9525;

/** Convert pixels to EMUs */
export const pixelsToEmus = (pixels: number): number => Math.round(pixels * EMU_PER_PIXEL);

/** Convert EMUs to pixels */
export const emusToPixels = (emus: number): number => Math.round(emus / EMU_PER_PIXEL);

/** Convert inches to EMUs */
export const inchesToEmus = (inches: number): number => Math.round(inches * 914400);

/** Convert EMUs to inches */
export const emusToInches = (emus: number): number => emus / 914400;

/** Convert points to EMUs (1 point = 12700 EMUs) */
export const pointsToEmus = (points: number): number => Math.round(points * 12700);

/** Convert EMUs to points */
export const emusToPoints = (emus: number): number => emus / 12700;

/** Convert percentage to thousandths of a percent (used in some OOXML attributes) */
export const percentToTHousandths = (percent: number): number => Math.round(percent * 1000);

/** Common slide size presets (in pixels at 96 DPI) */
export const SlideSizePreset = {
    WIDE: { width: 960, height: 540 },
    STANDARD_4X3: { width: 720, height: 540 },
    WIDESCREEN_16X10: { width: 960, height: 600 },
    WIDESCREEN_16X9: { width: 960, height: 540 },
} as const;
