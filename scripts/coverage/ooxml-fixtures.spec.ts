import { describe, expect, it } from "vite-plus/test";

import { buildOoxmlExecutableReport, executeOoxmlCase } from "./ooxml-executable";
import { OOXML_EXECUTABLE_REGISTRY } from "./ooxml-fixtures";

describe("executable OOXML fixture harness", () => {
  it("executes every registered OOXML fixture", () => {
    expect(OOXML_EXECUTABLE_REGISTRY.cases.length).toBe(36);
    for (const fixtureCase of OOXML_EXECUTABLE_REGISTRY.cases) {
      expect(() => executeOoxmlCase(fixtureCase)).not.toThrow();
    }
  });

  it("fully owns the executable PoC schemas", () => {
    const report = buildOoxmlExecutableReport([
      "w16cid",
      "w16cex",
      "w15",
      "w14",
      "dml-lc",
      "dml-pic",
      "dml-cdr",
      "dml-xdr",
    ]);
    expect(report.mode).toBe("executable");
    expect(report.elementOwnership).toBe("152/152");
    expect(report.attributeOwnership).toBe("178/178");
    expect(report.executedCases).toBe(36);
    expect(report.unknownFixtures).toEqual([]);
    expect(report.missingElementOwnership).toEqual([]);
    expect(report.missingAttributeOwnership).toEqual([]);
    expect(report.roundTripFailures).toEqual([]);
    expect(report.semanticFailures).toEqual([]);
  });
});
