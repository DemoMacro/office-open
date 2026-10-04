import { describe, expect, it } from "vite-plus/test";

import { generateOdp, OdpParseError, parseOdp } from "./index";

describe("ODP codec", () => {
  it("round-trips a presentation through canonical PresentationOptions", () => {
    const parsed = parseOdp(generateOdp({ slides: [{ notes: "ODP" }] }));
    expect(parsed.slides?.[0]?.notes).toBe("ODP");
  });

  it("wraps invalid packages in OdpParseError", () => {
    expect(() => parseOdp(new Uint8Array([1, 2, 3]))).toThrow(OdpParseError);
  });
});
