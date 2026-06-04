import { describe, expect, it } from "vite-plus/test";

import { Bibliography } from "./bibliography";

describe("Bibliography", () => {
  it("should expose a Relationships instance", () => {
    const bib = new Bibliography({ sources: [] });
    expect(bib.relationships).to.be.ok;
    expect(bib.relationships.relationshipCount).to.equal(0);
  });
});
