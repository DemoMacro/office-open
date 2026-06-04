import { describe, expect, it } from "vite-plus/test";

import { AbstractNumbering } from "./abstract-numbering";

describe("AbstractNumbering", () => {
  it("stores its ID at its .id property", () => {
    const abstractNumbering = new AbstractNumbering(5, []);
    expect(abstractNumbering.id).to.equal(5);
  });
});
