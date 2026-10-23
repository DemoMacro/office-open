import { strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import { OpcPackageError } from "./package-error";
import { ParsedArchive } from "./parser";
import { nativeZip, readCentralDirectory } from "./zip-native";

describe("native ZIP writer", () => {
  it("round-trips Unicode filenames", () => {
    const data = new TextEncoder().encode("audio");
    const archive = nativeZip({ "ppt/media/Café.wav": data });
    expect(unzipSync(archive)["ppt/media/Café.wav"]).toEqual(data);
    expect(readCentralDirectory(archive)[0]?.name).toBe("ppt/media/Café.wav");
  });
});

describe("ZIP CRC validation", () => {
  it("rejects corrupted STORE entries instead of returning their bytes", () => {
    const archive = zipSync({ "a.xml": strToU8("<root/>") }, { level: 0 });
    const entry = readCentralDirectory(archive)[0]!;
    archive[entry.dataStart] = archive[entry.dataStart]! ^ 1;
    const parsed = new ParsedArchive(archive);
    expectCorruptEntry(parsed, "a.xml");
  });

  it("rejects corrupted DEFLATE entries instead of retrying check-free", () => {
    const archive = zipSync({ "a.xml": strToU8("<root/>") }, { level: 6 });
    const entry = readCentralDirectory(archive)[0]!;
    const parsed = new ParsedArchive(archive, [{ ...entry, crc: entry.crc ^ 1 }]);
    expectCorruptEntry(parsed, "a.xml");
  });
});

function expectCorruptEntry(parsed: ParsedArchive, name: string): void {
  let cause: unknown;
  try {
    parsed.get(name);
  } catch (error) {
    cause = error;
  }
  expect(cause).toBeInstanceOf(OpcPackageError);
  expect((cause as OpcPackageError).reason).toBe("invalid-entry");
  expect((cause as OpcPackageError).cause).toMatchObject({
    message: `ZIP CRC-32 mismatch: ${name}`,
    reason: "crc-mismatch",
  });
}
