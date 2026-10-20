import { unzipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import { nativeZip, readCentralDirectory } from "./zip-native";

describe("native ZIP writer", () => {
  it("round-trips Unicode filenames", () => {
    const data = new TextEncoder().encode("audio");
    const archive = nativeZip({ "ppt/media/Café.wav": data });
    expect(unzipSync(archive)["ppt/media/Café.wav"]).toEqual(data);
    expect(readCentralDirectory(archive)[0]?.name).toBe("ppt/media/Café.wav");
  });
});
