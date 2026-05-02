import { describe, expect, it } from "vite-plus/test";

import {
    decimalNumber,
    eighthPointMeasureValue,
    hexColorValue,
    hpsMeasureValue,
    longHexNumber,
    measurementOrPercentValue,
    percentageValue,
    pointMeasureValue,
    positiveUniversalMeasureValue,
    shortHexNumber,
    signedHpsMeasureValue,
    signedTwipsMeasureValue,
    twipsMeasureValue,
    uCharHexNumber,
    universalMeasureValue,
    unsignedDecimalNumber,
    ThemeColor,
    ThemeFont,
} from "./values";

describe("decimalNumber", () => {
    it("should floor a positive float", () => {
        expect(decimalNumber(10.7)).toBe(10);
    });

    it("should floor a negative float", () => {
        expect(decimalNumber(-5.3)).toBe(-6);
    });

    it("should pass through an integer", () => {
        expect(decimalNumber(42)).toBe(42);
    });

    it("should handle zero", () => {
        expect(decimalNumber(0)).toBe(0);
    });

    it("should throw on NaN", () => {
        expect(() => decimalNumber(NaN)).toThrow("Invalid value 'NaN'");
    });
});

describe("unsignedDecimalNumber", () => {
    it("should floor a positive float", () => {
        expect(unsignedDecimalNumber(10.7)).toBe(10);
    });

    it("should throw on negative value", () => {
        expect(() => unsignedDecimalNumber(-5)).toThrow("Must be a positive integer");
    });

    it("should accept zero", () => {
        expect(unsignedDecimalNumber(0)).toBe(0);
    });
});

describe("longHexNumber", () => {
    it("should accept a valid 8-char hex string", () => {
        expect(longHexNumber("ABCD1234")).toBe("ABCD1234");
    });

    it("should accept lowercase", () => {
        expect(longHexNumber("abcd1234")).toBe("abcd1234");
    });

    it("should throw on wrong length", () => {
        expect(() => longHexNumber("ABC")).toThrow("Expected 8 digit hex value");
    });

    it("should throw on invalid chars", () => {
        expect(() => longHexNumber("GHIJ1234")).toThrow("Expected 8 digit hex value");
    });
});

describe("shortHexNumber", () => {
    it("should accept a valid 4-char hex string", () => {
        expect(shortHexNumber("AB12")).toBe("AB12");
    });

    it("should throw on wrong length", () => {
        expect(() => shortHexNumber("ABC")).toThrow("Expected 4 digit hex value");
    });
});

describe("uCharHexNumber", () => {
    it("should accept a valid 2-char hex string", () => {
        expect(uCharHexNumber("FF")).toBe("FF");
    });

    it("should throw on wrong length", () => {
        expect(() => uCharHexNumber("FFF")).toThrow("Expected 2 digit hex value");
    });
});

describe("hexColorValue", () => {
    it("should accept 'auto'", () => {
        expect(hexColorValue("auto")).toBe("auto");
    });

    it("should accept 6-char hex", () => {
        expect(hexColorValue("FF0000")).toBe("FF0000");
    });

    it("should strip # prefix", () => {
        expect(hexColorValue("#00FF00")).toBe("00FF00");
    });

    it("should throw on invalid hex", () => {
        expect(() => hexColorValue("GG0000")).toThrow();
    });

    it("should throw on wrong length hex", () => {
        expect(() => hexColorValue("FFF")).toThrow();
    });
});

describe("universalMeasureValue", () => {
    it("should normalize mm", () => {
        expect(universalMeasureValue("10.500mm")).toBe("10.5mm");
    });

    it("should normalize cm", () => {
        expect(universalMeasureValue("2cm")).toBe("2cm");
    });

    it("should normalize in", () => {
        expect(universalMeasureValue("1.5in")).toBe("1.5in");
    });

    it("should normalize pt", () => {
        expect(universalMeasureValue("12pt")).toBe("12pt");
    });

    it("should normalize pc", () => {
        expect(universalMeasureValue("3pc")).toBe("3pc");
    });

    it("should normalize pi", () => {
        expect(universalMeasureValue("1pi")).toBe("1pi");
    });

    it("should preserve negative values", () => {
        expect(universalMeasureValue("-5mm")).toBe("-5mm");
    });
});

describe("positiveUniversalMeasureValue", () => {
    it("should accept positive values", () => {
        expect(positiveUniversalMeasureValue("10mm")).toBe("10mm");
    });

    it("should throw on negative values", () => {
        expect(() => positiveUniversalMeasureValue("-5mm")).toThrow("Expected a positive number");
    });
});

describe("signedTwipsMeasureValue", () => {
    it("should pass through universal measure string", () => {
        expect(signedTwipsMeasureValue("10mm")).toBe("10mm");
    });

    it("should floor a numeric value", () => {
        expect(signedTwipsMeasureValue(1440.7)).toBe(1440);
    });

    it("should accept negative numbers", () => {
        expect(signedTwipsMeasureValue(-100)).toBe(-100);
    });
});

describe("hpsMeasureValue", () => {
    it("should pass through positive universal measure", () => {
        expect(hpsMeasureValue("12pt")).toBe("12pt");
    });

    it("should validate positive number", () => {
        expect(hpsMeasureValue(24)).toBe(24);
    });

    it("should throw on negative number", () => {
        expect(() => hpsMeasureValue(-1)).toThrow();
    });
});

describe("signedHpsMeasureValue", () => {
    it("should pass through universal measure", () => {
        expect(signedHpsMeasureValue("6pt")).toBe("6pt");
    });

    it("should floor a number", () => {
        expect(signedHpsMeasureValue(12.5)).toBe(12);
    });

    it("should accept negative numbers", () => {
        expect(signedHpsMeasureValue(-6)).toBe(-6);
    });
});

describe("twipsMeasureValue", () => {
    it("should pass through positive universal measure", () => {
        expect(twipsMeasureValue("25.4mm")).toBe("25.4mm");
    });

    it("should validate positive number", () => {
        expect(twipsMeasureValue(1440)).toBe(1440);
    });

    it("should throw on negative number", () => {
        expect(() => twipsMeasureValue(-1)).toThrow();
    });
});

describe("percentageValue", () => {
    it("should normalize percentage", () => {
        expect(percentageValue("50.000%")).toBe("50%");
    });

    it("should handle integer percentage", () => {
        expect(percentageValue("100%")).toBe("100%");
    });

    it("should handle negative percentage", () => {
        expect(percentageValue("-10.5%")).toBe("-10.5%");
    });
});

describe("measurementOrPercentValue", () => {
    it("should floor a number", () => {
        expect(measurementOrPercentValue(100.5)).toBe(100);
    });

    it("should normalize a percentage", () => {
        expect(measurementOrPercentValue("50%")).toBe("50%");
    });

    it("should normalize a universal measure", () => {
        expect(measurementOrPercentValue("10mm")).toBe("10mm");
    });
});

describe("eighthPointMeasureValue", () => {
    it("should validate positive integer", () => {
        expect(eighthPointMeasureValue(16)).toBe(16);
    });

    it("should floor float", () => {
        expect(eighthPointMeasureValue(16.7)).toBe(16);
    });

    it("should throw on negative", () => {
        expect(() => eighthPointMeasureValue(-1)).toThrow();
    });
});

describe("pointMeasureValue", () => {
    it("should validate positive integer", () => {
        expect(pointMeasureValue(12)).toBe(12);
    });

    it("should throw on negative", () => {
        expect(() => pointMeasureValue(-1)).toThrow();
    });
});

describe("ThemeColor", () => {
    it("should have all 17 ST_ThemeColor values", () => {
        const values = Object.values(ThemeColor);
        expect(values).toHaveLength(17);
    });

    it("should include required values per XSD", () => {
        expect(ThemeColor.DARK1).toBe("dark1");
        expect(ThemeColor.LIGHT1).toBe("light1");
        expect(ThemeColor.ACCENT1).toBe("accent1");
        expect(ThemeColor.HYPERLINK).toBe("hyperlink");
        expect(ThemeColor.FOLLOWED_HYPERLINK).toBe("followedHyperlink");
        expect(ThemeColor.NONE).toBe("none");
        expect(ThemeColor.BACKGROUND1).toBe("background1");
        expect(ThemeColor.TEXT1).toBe("text1");
    });
});

describe("ThemeFont", () => {
    it("should have all 8 ST_Theme values", () => {
        const values = Object.values(ThemeFont);
        expect(values).toHaveLength(8);
    });

    it("should include required values per XSD", () => {
        expect(ThemeFont.MAJOR_ASCII).toBe("majorAscii");
        expect(ThemeFont.MAJOR_H_ANSI).toBe("majorHAnsi");
        expect(ThemeFont.MAJOR_EAST_ASIA).toBe("majorEastAsia");
        expect(ThemeFont.MAJOR_BIDI).toBe("majorBidi");
        expect(ThemeFont.MINOR_ASCII).toBe("minorAscii");
        expect(ThemeFont.MINOR_H_ANSI).toBe("minorHAnsi");
        expect(ThemeFont.MINOR_EAST_ASIA).toBe("minorEastAsia");
        expect(ThemeFont.MINOR_BIDI).toBe("minorBidi");
    });
});
