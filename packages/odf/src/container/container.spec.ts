import { parse } from "@office-open/xml";
import { strFromU8, strToU8, unzipSync, zipSync } from "fflate";
import { describe, expect, it } from "vite-plus/test";

import {
  generateOcf,
  compareOcfSemantic,
  manifestOptionsXml,
  manifestXml,
  OcfManifestError,
  ODF_ELEMENT_NAMES,
  parseManifestOptions,
  parseOdfNodes,
  readOcf,
  serializeOdfNodes,
  type OdfXmlNode,
} from "../index";

describe("OCF runtime", () => {
  it("contains every unique element in the ODF and manifest schemas", () => {
    expect(ODF_ELEMENT_NAMES).toHaveLength(610);
    expect(new Set(ODF_ELEMENT_NAMES).size).toBe(610);
  });

  it("round-trips generic ODF nodes and preserves foreign namespaces", () => {
    const nodes: OdfXmlNode[] = [
      {
        name: "style:style",
        attributes: { "style:name": "List", "style:family": "paragraph" },
        children: [
          { name: "style:paragraph-properties", attributes: { "fo:text-align": "start" } },
        ],
      },
      { name: "text:list", children: [{ name: "text:list-item", children: ["Item"] }] },
      { name: "draw:frame", attributes: { "draw:name": "Frame" } },
      { name: "table:named-range", attributes: { "table:name": "Total" } },
      { name: "number:number-style", attributes: { "style:name": "N1" } },
      { name: "chart:chart", attributes: { "chart:class": "chart:bar" } },
      { name: "anim:par", children: [{ name: "anim:set" }] },
      { name: "config:config-item", attributes: { "config:name": "ShowGrid" } },
      { name: "db:data-source", attributes: { "db:name": "Library" } },
      { name: "form:form", attributes: { "form:name": "Editor" } },
      { name: "manifest:file-entry", attributes: { "manifest:full-path": "/" } },
      { name: "unknown:element", attributes: { hidden: "true" } },
    ];
    const serialized = serializeOdfNodes(nodes).join("");
    const document = parse(
      `<root xmlns:style="urn:1" xmlns:text="urn:2" xmlns:draw="urn:3" xmlns:table="urn:4" xmlns:number="urn:5" xmlns:chart="urn:6" xmlns:anim="urn:7" xmlns:config="urn:8" xmlns:db="urn:9" xmlns:form="urn:10" xmlns:manifest="urn:11" xmlns:unknown="urn:12">${serialized}</root>`,
      { ignoreDeclaration: true },
    );
    expect(parseOdfNodes(document.elements?.[0])).toEqual(nodes);
  });

  it("declares OCF subdirectories and packs binary resources", () => {
    const mimeType = "application/vnd.oasis.opendocument.chart";
    const files = {
      "Object 1/content.xml": "<chart:chart/>",
      "Thumbnails/thumbnail.png": new Uint8Array([1, 2, 3]),
    };
    const archive = generateOcf(mimeType, files);
    expect(archive.subarray(0, 2)).toEqual(new Uint8Array([0x50, 0x4b]));
    expect(manifestXml(mimeType, files)).toContain('manifest:full-path="Object 1/"');
    expect(manifestXml(mimeType, files)).toContain('manifest:full-path="Thumbnails/"');
  });

  it("opens a manifest-identified package without a mimetype entry", () => {
    const archive = zipSync({
      "content.xml": strToU8(
        '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>',
      ),
      "META-INF/manifest.xml": strToU8(
        '<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>',
      ),
    });
    expect(() => readOcf(archive, "application/vnd.oasis.opendocument.text")).not.toThrow();
  });

  it("ignores stale manifest declarations for absent package paths", () => {
    const archive = zipSync({
      mimetype: strToU8("application/vnd.oasis.opendocument.text"),
      "content.xml": strToU8(
        '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>',
      ),
      "META-INF/manifest.xml": strToU8(
        '<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="META-INF/manifest.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="Thumbnails/thumbnail.png" manifest:media-type="image/png"/></manifest:manifest>',
      ),
    });
    const parsed = readOcf(archive, "application/vnd.oasis.opendocument.text");
    expect(parsed.manifest.entries.map((entry) => entry.fullPath)).not.toContain(
      "Thumbnails/thumbnail.png",
    );
  });

  it("reads a typed manifest and preserves encryption metadata", () => {
    const xml = `<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.chart"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"><manifest:encryption-data manifest:checksum-type="SHA1/1K" manifest:checksum="aa=="><manifest:algorithm manifest:algorithm-name="Blowfish CFB" manifest:initialisation-vector="bb=="/><manifest:key-derivation manifest:key-derivation-name="PBKDF2" manifest:salt="cc==" manifest:iteration-count="1024" manifest:key-size="16"/></manifest:encryption-data></manifest:file-entry></manifest:manifest>`;
    const parsed = parseManifestOptions(xml);
    expect(parsed.version).toBe("1.3");
    expect(parsed.entries[1]?.encryptionData?.keyDerivation).toEqual({
      name: "PBKDF2",
      salt: "cc==",
      iterationCount: 1024,
      keySize: 16,
    });
    expect(manifestOptionsXml(parsed)).toContain('manifest:iteration-count="1024"');
  });

  it("round-trips encrypted keys and file-entry encryption through one typed model", () => {
    const xml = `<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:encrypted-key><manifest:encryption-method manifest:PGPAlgorithm="RSA"/><manifest:keyinfo><manifest:PGPData><manifest:PGPKeyID>MQ==</manifest:PGPKeyID><manifest:PGPKeyPacket>Ag==</manifest:PGPKeyPacket></manifest:PGPData></manifest:keyinfo><manifest:CipherData><manifest:CipherValue>Aw==</manifest:CipherValue></manifest:CipherData></manifest:encrypted-key><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.chart"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"><manifest:encryption-data manifest:checksum-type="SHA1/1K" manifest:checksum="aa=="><manifest:algorithm manifest:algorithm-name="Blowfish CFB" manifest:initialisation-vector="bb=="/><manifest:key-derivation manifest:key-derivation-name="PBKDF2" manifest:salt="cc==" manifest:iteration-count="1024" manifest:key-size="16"/></manifest:encryption-data></manifest:file-entry></manifest:manifest>`;
    const parsed = parseManifestOptions(xml);
    expect(parsed.encryptedKeys).toEqual([
      { algorithm: "RSA", keyId: "MQ==", keyPacket: "Ag==", cipherValue: "Aw==" },
    ]);
    expect(parseManifestOptions(manifestOptionsXml(parsed))).toEqual(parsed);

    const archive = generateOcf(
      "application/vnd.oasis.opendocument.chart",
      { "content.xml": "<a/>" },
      {},
      parsed,
    );
    const reread = readOcf(archive, "application/vnd.oasis.opendocument.chart").manifest;
    const packaged = {
      ...parsed,
      entries: [
        ...parsed.entries,
        { fullPath: "mimetype", mediaType: "application/vnd.oasis.opendocument.chart" },
      ],
    };
    expect(reread).toEqual(packaged);
    expect(
      readOcf(
        generateOcf(
          "application/vnd.oasis.opendocument.chart",
          { "content.xml": "<a/>" },
          {},
          reread,
        ),
        "application/vnd.oasis.opendocument.chart",
      ).manifest,
    ).toEqual(packaged);
  });

  it("keeps supplied manifest entries and reports duplicate paths structurally", () => {
    const files = { "content.xml": "<a/>", "Extra/data.bin": new Uint8Array([1]) };
    const packageManifest = {
      version: "1.3",
      entries: [
        { fullPath: "/", version: "1.3", mediaType: "application/vnd.oasis.opendocument.chart" },
        { fullPath: "Extra/data.bin", mediaType: "custom/source-type" },
      ],
    };
    const manifest = strFromU8(
      unzipSync(
        generateOcf("application/vnd.oasis.opendocument.chart", files, {}, packageManifest),
      )["META-INF/manifest.xml"]!,
    );
    expect(manifest).toContain(
      'manifest:full-path="Extra/data.bin" manifest:media-type="custom/source-type"',
    );
    expect(
      compareOcfSemantic(
        generateOcf("application/vnd.oasis.opendocument.chart", files, {}, packageManifest),
        generateOcf("application/vnd.oasis.opendocument.chart", files, {}, packageManifest),
      ),
    ).toEqual([]);
    let error: unknown;
    try {
      generateOcf(
        "application/vnd.oasis.opendocument.chart",
        files,
        {},
        {
          ...packageManifest,
          entries: [...packageManifest.entries, packageManifest.entries[1]!],
        },
      );
    } catch (cause) {
      error = cause;
    }
    expect(error).toBeInstanceOf(OcfManifestError);
    expect(error).toHaveProperty("fullPath", "Extra/data.bin");
  });

  it("keeps manifest media types over derived types and rejects missing declared paths", () => {
    const files = { "Extra/data.bin": new Uint8Array([1]) };
    const packageManifest = {
      version: "1.3",
      entries: [
        { fullPath: "/", version: "1.3", mediaType: "application/vnd.oasis.opendocument.chart" },
        { fullPath: "Extra/", mediaType: "application/vnd.sun.xml.ui.configuration" },
        { fullPath: "Extra/data.bin", mediaType: "custom/source-type" },
      ],
    };
    const manifest = manifestXml(
      "application/vnd.oasis.opendocument.chart",
      files,
      { "Extra/data.bin": "application/binary" },
      packageManifest,
    );
    expect(manifest).toContain(
      'manifest:full-path="Extra/data.bin" manifest:media-type="custom/source-type"',
    );

    let error: unknown;
    try {
      manifestXml(
        "application/vnd.oasis.opendocument.chart",
        files,
        {},
        {
          ...packageManifest,
          entries: [
            ...packageManifest.entries,
            { fullPath: "Missing/data.bin", mediaType: "application/binary" },
          ],
        },
      );
    } catch (cause) {
      error = cause;
    }
    expect(error).toBeInstanceOf(OcfManifestError);
    expect(error).toHaveProperty("fullPath", "Missing/data.bin");
  });

  it("rejects a manifest root media type that differs from the package mimetype", () => {
    expect(() =>
      manifestXml(
        "application/vnd.oasis.opendocument.chart",
        { "content.xml": "<a/>" },
        {},
        {
          version: "1.3",
          entries: [
            { fullPath: "/", mediaType: "application/vnd.oasis.opendocument.text" },
            { fullPath: "content.xml", mediaType: "text/xml" },
          ],
        },
      ),
    ).toThrow(OcfManifestError);
  });

  it("rejects unknown manifest children instead of silently skipping them", () => {
    const xml = `<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" manifest:version="1.3"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.chart"/><manifest:unknown/></manifest:manifest>`;
    expect(() => parseManifestOptions(xml)).toThrow(OcfManifestError);
  });
});
