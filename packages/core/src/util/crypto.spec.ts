import { afterEach, describe, expect, it, vi } from "vite-plus/test";

import { randomBytes } from "./crypto";

describe("randomBytes", () => {
  afterEach(() => {
    vi.unstubAllGlobals();
  });

  it("fills bytes from Web Crypto", () => {
    const bytes = randomBytes(16);
    expect(bytes).toBeInstanceOf(Uint8Array);
    expect(bytes).toHaveLength(16);
  });

  it("fails closed when Web Crypto is unavailable", () => {
    vi.stubGlobal("crypto", undefined);
    expect(() => randomBytes(1)).toThrow(TypeError);
  });
});
