import { describe, expect, it } from "vitest";

import { decryptLegacyRc4, deriveLegacyRc4Key } from "./legacy-encryption";

describe("legacy Office encryption", () => {
  it("derives the documented legacy RC4 key", () => {
    const salt = Uint8Array.from([
      0xe8, 0x77, 0x2c, 0x1d, 0x91, 0xc5, 0x6a, 0x37, 0x96, 0x47, 0x61, 0xb2, 0x80, 0x18, 0x32,
      0x17,
    ]);
    expect(deriveLegacyRc4Key("password1", salt, 0)).toEqual(
      Uint8Array.from([
        0x20, 0xbf, 0x32, 0xdd, 0xf5, 0x40, 0x85, 0x8c, 0x51, 0x37, 0x44, 0xaf, 0x0f, 0x24, 0xe0,
        0x3c,
      ]),
    );
  });

  it("round-trips a block-restarted RC4 payload", () => {
    const salt = new Uint8Array(16).fill(7);
    const data = new Uint8Array(600).map((_, index) => index & 0xff);
    const encrypted = decryptLegacyRc4(data, "password", salt);
    expect(decryptLegacyRc4(encrypted, "password", salt)).toEqual(data);
  });
});
