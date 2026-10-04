import { readFileSync } from "node:fs";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { inflateRawSync } from "node:zlib";

import { expect, describe, it } from "vite-plus/test";

import { LEGACY_CAPABILITY_REGISTRY } from "./legacy-capability-registry";
import { ODF_CODEC_REGISTRY } from "./odf-codec-registry";

const __dirname = path.dirname(fileURLToPath(import.meta.url));

describe("coverage registry gates", () => {
  it("requires every ODF mapper to be executable and non-generic", () => {
    expect(ODF_CODEC_REGISTRY.length).toBeGreaterThan(0);
    for (const entry of ODF_CODEC_REGISTRY) {
      expect(entry.classification).not.toBe("unsupported");
      expect(entry.classification).not.toBe("generic-only");
      expect(entry.schemaElements?.length).toBeGreaterThan(0);
      expect(entry.owner.export).toBeTruthy();
      expect(entry.roundTrip?.length).toBe(2);
      expect(entry.testId).toBeTruthy();
      expect(entry.fixtureKey).toBeTruthy();
      for (const negative of entry.negativeCapabilities ?? []) {
        expect(negative.testId).toBeTruthy();
      }
    }
  });

  it("requires legacy capabilities to name implementation, tests, and fixtures", () => {
    expect(LEGACY_CAPABILITY_REGISTRY.length).toBe(4);
    for (const format of LEGACY_CAPABILITY_REGISTRY) {
      expect(format.capabilities.length).toBeGreaterThan(0);
      for (const capability of format.capabilities) {
        expect(capability.supported).toBe(true);
        expect(capability.owner.export).toBeTruthy();
        expect(["container", "record", "stream", "mapper"]).toContain(capability.mapper.layer);
        expect(capability.mapper.export).toBeTruthy();
        expect(capability.assertion).toBeTruthy();
        expect(capability.testId).toBeTruthy();
        expect(capability.fixture.key).toBeTruthy();
        expect(capability.fixture.generator).toBeTruthy();
      }
    }
  });

  it("uses non-empty deterministic cross-format converted fixtures", () => {
    for (const name of ["doc.b64z", "ppt.b64z", "xls.b64z"]) {
      const encoded = readFileSync(path.resolve(__dirname, "../fixtures", name), "utf8").trim();
      const data = inflateRawSync(Buffer.from(encoded, "base64"));
      expect(data.byteLength).toBeGreaterThan(0);
      expect(data[0]).toBe(0xd0);
      expect(data[1]).toBe(0xcf);
    }
  });

  it("fails empty or missing corpora instead of skipping them", () => {
    const corpus = readFileSync(path.resolve(__dirname, "../corpus-check.ts"), "utf8");
    const legacyCorpus = readFileSync(path.resolve(__dirname, "../legacy-corpus-check.ts"), "utf8");
    expect(corpus).toContain("corpus is incomplete");
    expect(corpus).not.toContain("(skipped)");
    expect(legacyCorpus).toContain("corpus is incomplete");
    expect(legacyCorpus).toContain("empty corpus");
  });
});
