import { strFromU8, strToU8, unzipSync, zipSync } from "@office-open/core";
import { parse } from "@office-open/xml";
import { describe, expect, it } from "vite-plus/test";

import {
  generateOcf,
  compareOcfSemantic,
  manifestOptionsXml,
  manifestXml,
  OcfManifestError,
  ODF_ELEMENT_NAMES,
  OdfXmlError,
  parseManifestOptions,
  parseOdfNodes,
  readOcf,
  readXml,
  stringifyOdfNodes,
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
    const serialized = stringifyOdfNodes(nodes).join("");
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

  it("opens a manifest-identified package without a mimetype entry", async () => {
    const archive = zipSync({
      "content.xml": strToU8(
        '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>',
      ),
      "META-INF/manifest.xml": strToU8(
        '<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>',
      ),
    });
    await expect(
      readOcf(archive, "application/vnd.oasis.opendocument.text"),
    ).resolves.toBeDefined();
  });

  it("accepts any expected family mimetype, including templates", async () => {
    const mimeType = "application/vnd.oasis.opendocument.text";
    const templateMimeType = "application/vnd.oasis.opendocument.text-template";
    const archive = generateOcf(templateMimeType, {
      "content.xml":
        '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>',
    });
    await expect(readOcf(archive, [mimeType, templateMimeType])).resolves.toBeDefined();
  });

  it("normalizes browser binary inputs before reading an OCF package", async () => {
    const mimeType = "application/vnd.oasis.opendocument.text";
    const archive = generateOcf(mimeType, { "content.xml": "<office:document-content/>" });
    const buffer = new ArrayBuffer(archive.byteLength);
    new Uint8Array(buffer).set(archive);
    const parsed = await readOcf(new Blob([buffer]), mimeType);
    expect(parsed.files["content.xml"]).toContain("<office:document-content/>");
  });

  it("normalizes ODF element prefixes by namespace URI", () => {
    const root = readXml(
      {
        "content.xml":
          '<o:document-content xmlns:o="urn:oasis:names:tc:opendocument:xmlns:office:1.0"><o:body/></o:document-content>',
      },
      "content.xml",
    );

    expect(root.name).toBe("office:document-content");
    expect(root.elements?.[0]?.name).toBe("office:body");
  });

  it("rejects XML parts with multiple roots or text outside the root", () => {
    for (const xml of ["<root/><extra/>", "text<root/>", "<root/>text"]) {
      expect(() => readXml({ "content.xml": xml }, "content.xml")).toThrow(OdfXmlError);
    }
    expect(() =>
      readXml({ "content.xml": "<!-- before -->\n  <root/>\n<!-- after -->" }, "content.xml"),
    ).not.toThrow();
  });

  it("rejects manifests with multiple roots or text outside the root", () => {
    const root =
      '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/></manifest:manifest>';
    for (const xml of [`${root}<extra/>`, `text${root}`, `${root}text`]) {
      expect(() => parseManifestOptions(xml)).toThrow(OcfManifestError);
    }
  });

  it("round-trips explicit empty ZIP directories", async () => {
    const mimeType = "application/vnd.oasis.opendocument.text";
    const packageManifest = {
      version: "1.3",
      entries: [
        { fullPath: "/", version: "1.3", mediaType: mimeType },
        { fullPath: "content.xml", mediaType: "text/xml" },
        { fullPath: "Configurations2/", mediaType: "application/binary" },
      ],
    };
    const source = generateOcf(
      mimeType,
      {
        "content.xml":
          '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>',
      },
      {},
      packageManifest,
    );
    expect(Object.keys(unzipSync(source))).toContain("Configurations2/");
    const parsed = await readOcf(source, mimeType);
    expect(parsed.manifest.entries.map((entry) => entry.fullPath)).toContain("Configurations2/");
    expect(
      Object.keys(unzipSync(generateOcf(mimeType, parsed.files, {}, parsed.manifest))),
    ).toContain("Configurations2/");
  });

  it("rejects stale manifest declarations for absent package paths", async () => {
    const archive = zipSync({
      mimetype: strToU8("application/vnd.oasis.opendocument.text"),
      "content.xml": strToU8(
        '<office:document-content xmlns:office="urn:oasis:names:tc:opendocument:xmlns:office:1.0"/>',
      ),
      "META-INF/manifest.xml": strToU8(
        '<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="META-INF/manifest.xml" manifest:media-type="text/xml"/><manifest:file-entry manifest:full-path="required-package-member.bin" manifest:media-type="application/binary"/></manifest:manifest>',
      ),
    });
    await expect(readOcf(archive, "application/vnd.oasis.opendocument.text")).rejects.toThrow(
      "ODF manifest declares missing package member: required-package-member.bin",
    );
  });

  it("rejects ZIP members whose declared uncompressed size exceeds the ODF limit", async () => {
    const archive = generateOcf("application/vnd.oasis.opendocument.text", {
      "content.xml": "<a/>",
    });
    for (let offset = 0; offset < archive.length - 46; offset += 1) {
      if (
        archive[offset] !== 0x50 ||
        archive[offset + 1] !== 0x4b ||
        archive[offset + 2] !== 0x01 ||
        archive[offset + 3] !== 0x02
      ) {
        continue;
      }
      const view = new DataView(archive.buffer, archive.byteOffset, archive.byteLength);
      const nameLength = view.getUint16(offset + 28, true);
      const name = strFromU8(archive.subarray(offset + 46, offset + 46 + nameLength));
      if (name === "content.xml") {
        view.setUint32(offset + 24, 256 * 1024 * 1024 + 1, true);
        await expect(
          readOcf(archive, "application/vnd.oasis.opendocument.text"),
        ).rejects.toMatchObject({
          cause: expect.objectContaining({ reason: "size-limit" }),
        });
        return;
      }
      offset +=
        45 + nameLength + view.getUint16(offset + 30, true) + view.getUint16(offset + 32, true);
    }
    throw new Error("content.xml central directory entry was not found");
  });

  it("accepts and retains auxiliary package members omitted by legacy manifests", async () => {
    const archive = zipSync({
      mimetype: strToU8("application/vnd.oasis.opendocument.text"),
      "content.xml": strToU8(
        '<office:document-content xmlns:office="urn:oasis-names:tc:opendocument:xmlns:office:1.0"/>',
      ),
      "ObjectReplacements/Object 1": new Uint8Array([1, 2, 3]),
      "META-INF/manifest.xml": strToU8(
        '<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis-names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>',
      ),
    });
    const parsed = await readOcf(archive, "application/vnd.oasis.opendocument.text");
    expect(parsed.binaries["ObjectReplacements/Object 1"]).toEqual(new Uint8Array([1, 2, 3]));
    expect(parsed.manifest.entries).toMatchObject([
      { fullPath: "/" },
      { fullPath: "content.xml" },
      { fullPath: "ObjectReplacements/Object 1", mediaType: "application/binary" },
    ]);
  });

  it("rejects duplicate and invalid manifest paths", async () => {
    const manifest = (path: string, duplicate = false) =>
      `<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="${path}" manifest:media-type="text/xml"/>${duplicate ? `<manifest:file-entry manifest:full-path="${path}" manifest:media-type="text/xml"/>` : ""}</manifest:manifest>`;
    const create = (path: string, duplicate = false) =>
      zipSync({
        mimetype: strToU8("application/vnd.oasis.opendocument.text"),
        "content.xml": strToU8("<office:document-content/>"),
        "META-INF/manifest.xml": strToU8(manifest(path, duplicate)),
      });
    await expect(
      readOcf(create("content.xml", true), "application/vnd.oasis.opendocument.text"),
    ).rejects.toThrow("Duplicate ODF manifest path: content.xml");
    await expect(
      readOcf(create("../content.xml"), "application/vnd.oasis.opendocument.text"),
    ).rejects.toThrow("Invalid ODF manifest path: ../content.xml");
  });

  it.each([
    "../outside.xml",
    "/absolute.xml",
    "C:/outside.xml",
    "mimetype",
    "META-INF/manifest.xml",
  ])("rejects unsafe or reserved generated package member %s", (path) => {
    expect(() =>
      generateOcf("application/vnd.oasis.opendocument.text", { [path]: "<a/>" }),
    ).toThrow("Invalid ODF package member path");
  });

  it("rejects unsafe package-manifest paths during generation", () => {
    expect(() =>
      generateOcf(
        "application/vnd.oasis.opendocument.text",
        { "content.xml": "<a/>" },
        {},
        { version: "1.3", entries: [{ fullPath: "../outside/", mediaType: "" }] },
      ),
    ).toThrow("Invalid ODF manifest path: ../outside/");
  });

  it("rejects file and directory path conflicts during generation", () => {
    const mimeType = "application/vnd.oasis.opendocument.text";
    expect(() =>
      generateOcf(
        mimeType,
        { "Object 1": "file" },
        {},
        { version: "1.3", entries: [{ fullPath: "Object 1/", mediaType: "" }] },
      ),
    ).toThrow("ODF package file/directory path conflict: Object 1");
    expect(() =>
      generateOcf(mimeType, { "Object 1": "file", "Object 1/content.xml": "<a/>" }),
    ).toThrow("ODF package file/directory path conflict: Object 1");
    for (const directory of ["mimetype/", "META-INF/manifest.xml/"]) {
      expect(() =>
        generateOcf(
          mimeType,
          { "content.xml": "<a/>" },
          {},
          {
            version: "1.3",
            entries: [{ fullPath: directory, mediaType: "application/binary" }],
          },
        ),
      ).toThrow(/ODF package file\/directory path conflict/u);
    }
  });

  it("rejects file and directory path conflicts when reading", async () => {
    const mimeType = "application/vnd.oasis.opendocument.text";
    const archive = zipSync({
      mimetype: strToU8(mimeType),
      "Object 1": strToU8("file"),
      "Object 1/": new Uint8Array(),
      "META-INF/manifest.xml": strToU8(
        `<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="${mimeType}"/><manifest:file-entry manifest:full-path="Object 1" manifest:media-type="application/binary"/><manifest:file-entry manifest:full-path="Object 1/" manifest:media-type="application/binary"/></manifest:manifest>`,
      ),
    });
    await expect(readOcf(archive, mimeType)).rejects.toThrow(
      "ODF package file/directory path conflict: Object 1",
    );
  });

  it("rejects disagreement between ZIP mimetype and manifest root media type", async () => {
    const archive = unzipSync(
      generateOcf("application/vnd.oasis.opendocument.text", { "content.xml": "<a/>" }),
    );
    archive["META-INF/manifest.xml"] = strToU8(
      strFromU8(archive["META-INF/manifest.xml"]!).replace(
        /(manifest:full-path="\/"[^>]*manifest:media-type=")[^"]+"/,
        '$1application/vnd.oasis.opendocument.spreadsheet"',
      ),
    );

    await expect(
      readOcf(zipSync(archive), "application/vnd.oasis.opendocument.text"),
    ).rejects.toThrow("Unexpected ODF MIME type");
  });

  it("rejects unknown manifest-namespaced attributes", async () => {
    const archive = zipSync({
      mimetype: strToU8("application/vnd.oasis.opendocument.text"),
      "content.xml": strToU8("<office:document-content/>"),
      "META-INF/manifest.xml": strToU8(
        '<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text" manifest:unknown="1"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"/></manifest:manifest>',
      ),
    });
    await expect(readOcf(archive, "application/vnd.oasis.opendocument.text")).rejects.toThrow(
      "Unknown manifest attribute @manifest:unknown: manifest:file-entry",
    );
  });

  it("rejects foreign attributes not allowed by the manifest schema", () => {
    const xml =
      '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" xmlns:x="urn:extension"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text" x:flag="true"/></manifest:manifest>';

    expect(() => parseManifestOptions(xml)).toThrow(
      "Unknown manifest attribute @x:flag: manifest:file-entry",
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

  it("rejects duplicate file encryption and encryption parameter elements", () => {
    const root =
      '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/>';
    const entry =
      '<manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml">';
    const encryption =
      '<manifest:encryption-data manifest:checksum-type="SHA1/1K" manifest:checksum="aa=="><manifest:algorithm manifest:algorithm-name="Blowfish CFB" manifest:initialisation-vector="bb=="/><manifest:key-derivation manifest:key-derivation-name="PBKDF2"/></manifest:encryption-data>';
    const close = "</manifest:file-entry></manifest:manifest>";

    expect(() => parseManifestOptions(root + entry + encryption + encryption + close)).toThrow(
      /Duplicate manifest encryption-data/u,
    );
    expect(() =>
      parseManifestOptions(
        root +
          entry +
          '<manifest:encryption-data manifest:checksum-type="SHA1/1K" manifest:checksum="aa=="><manifest:algorithm manifest:algorithm-name="Blowfish CFB" manifest:initialisation-vector="bb=="/><manifest:algorithm manifest:algorithm-name="AES" manifest:initialisation-vector="cc=="/><manifest:key-derivation manifest:key-derivation-name="PBKDF2"/></manifest:encryption-data>' +
          close,
      ),
    ).toThrow(/Invalid manifest encryption-data/u);
  });

  it("preserves mixed foreign manifest content order", () => {
    const parsed = parseManifestOptions(
      '<?xml version="1.0"?><manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" xmlns:x="urn:extension"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"><manifest:encryption-data manifest:checksum-type="SHA1/1K" manifest:checksum="aa=="><manifest:algorithm manifest:algorithm-name="Blowfish CFB" manifest:initialisation-vector="bb=="><x:extension>before<![CDATA[middle]]><x:child/>after</x:extension></manifest:algorithm><manifest:key-derivation manifest:key-derivation-name="PBKDF2" manifest:salt="cc==" manifest:iteration-count="1024" manifest:key-size="16"/></manifest:encryption-data></manifest:file-entry></manifest:manifest>',
    );
    const output = manifestOptionsXml(parsed);

    expect(output.indexOf("before")).toBeLessThan(output.indexOf("<x:child"));
    expect(output.indexOf("<x:child")).toBeLessThan(output.indexOf("after"));
    expect(output.indexOf("before")).toBeLessThan(output.indexOf("<![CDATA[middle]]>"));
    expect(output.indexOf("<![CDATA[middle]]>")).toBeLessThan(output.indexOf("<x:child"));
    expect(output).toContain('xmlns:x="urn:extension"');
    expect(() => parse(output)).not.toThrow();
  });

  it("preserves attributes inside manifest algorithm extensions", () => {
    const parsed = parseManifestOptions(
      '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" xmlns:x="urn:extension"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"><manifest:encryption-data manifest:checksum-type="SHA1/1K" manifest:checksum="aa=="><manifest:algorithm manifest:algorithm-name="Blowfish CFB" manifest:initialisation-vector="bb=="><x:extension x:flag="value"/></manifest:algorithm><manifest:key-derivation manifest:key-derivation-name="PBKDF2"/></manifest:encryption-data></manifest:file-entry></manifest:manifest>',
    );
    const output = manifestOptionsXml(parsed);

    expect(output).toContain('<x:extension x:flag="value"');
    expect(() => parse(output)).not.toThrow();
  });

  it("rejects extension attributes on the manifest algorithm element", () => {
    const xml =
      '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" xmlns:x="urn:extension"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml"><manifest:encryption-data manifest:checksum-type="SHA1/1K" manifest:checksum="aa=="><manifest:algorithm manifest:algorithm-name="Blowfish CFB" manifest:initialisation-vector="bb==" x:flag="value"/><manifest:key-derivation manifest:key-derivation-name="PBKDF2"/></manifest:encryption-data></manifest:file-entry></manifest:manifest>';

    expect(() => parseManifestOptions(xml)).toThrow("Unknown manifest attribute @x:flag");
  });

  it("rejects non-whitespace text at the manifest root", () => {
    const xml =
      '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">unexpected<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/></manifest:manifest>';

    expect(() => parseManifestOptions(xml)).toThrow("unexpected text in manifest root");
  });

  it("rejects foreign elements not allowed by the manifest schema", () => {
    const xml =
      '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0" xmlns:x="urn:extension"><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/><x:extension/></manifest:manifest>';

    expect(() => parseManifestOptions(xml)).toThrow("Unknown manifest element: x:extension");
  });

  it("rejects encrypted-key entries missing required schema children", () => {
    const xml =
      '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"><manifest:encrypted-key><manifest:keyinfo><manifest:PGPData><manifest:PGPKeyPacket>Ag==</manifest:PGPKeyPacket></manifest:PGPData></manifest:keyinfo><manifest:CipherData><manifest:CipherValue>Aw==</manifest:CipherValue></manifest:CipherData></manifest:encrypted-key><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/></manifest:manifest>';
    expect(() => parseManifestOptions(xml)).toThrow(/encrypted-key structure/u);
  });

  it("rejects duplicate and incorrectly nested encrypted-key values", () => {
    const root =
      '<manifest:manifest xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0">';
    const suffix =
      '<manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/></manifest:manifest>';
    const key = (pgp: string, cipher: string) =>
      `<manifest:encrypted-key><manifest:keyinfo><manifest:PGPData>${pgp}</manifest:PGPData></manifest:keyinfo><manifest:CipherData>${cipher}</manifest:CipherData></manifest:encrypted-key>`;
    const id = "<manifest:PGPKeyID>MQ==</manifest:PGPKeyID>";
    const cipherValue = "<manifest:CipherValue>Aw==</manifest:CipherValue>";
    for (const entry of [
      key(`${id}${id}`, cipherValue),
      key(`<manifest:wrapper>${id}</manifest:wrapper>`, cipherValue),
    ]) {
      expect(() => parseManifestOptions(`${root}${entry}${suffix}`)).toThrow(
        /encrypted-key structure/u,
      );
    }
  });

  it("round-trips encrypted keys and file-entry encryption through one typed model", async () => {
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
    const reread = (await readOcf(archive, "application/vnd.oasis.opendocument.chart")).manifest;
    const packaged = {
      ...parsed,
      entries: [
        ...parsed.entries,
        { fullPath: "mimetype", mediaType: "application/vnd.oasis.opendocument.chart" },
      ],
    };
    expect(reread).toEqual(packaged);
    const reparsed = await readOcf(
      generateOcf(
        "application/vnd.oasis.opendocument.chart",
        { "content.xml": "<a/>" },
        {},
        reread,
      ),
      "application/vnd.oasis.opendocument.chart",
    );
    expect(reparsed.manifest).toEqual(packaged);
  });

  it("includes manifest encryption fields in the OCF semantic digest", () => {
    const mimeType = "application/vnd.oasis.opendocument.text";
    const createPackage = (iterationCount: number) =>
      generateOcf(
        mimeType,
        { "content.xml": "<office:document-content/>" },
        {},
        {
          version: "1.3",
          entries: [
            { fullPath: "/", mediaType: mimeType },
            {
              fullPath: "content.xml",
              mediaType: "text/xml",
              encryptionData: {
                checksumType: "SHA1/1K",
                checksum: "aa==",
                algorithm: { name: "Blowfish CFB", initialisationVector: "bb==" },
                keyDerivation: { name: "PBKDF2", salt: "cc==", iterationCount, keySize: 16 },
              },
            },
          ],
          encryptedKeys: [{ algorithm: "RSA", keyId: "MQ==", cipherValue: "Aw==" }],
        },
      );

    expect(compareOcfSemantic(createPackage(1024), createPackage(2048))).toEqual([
      'entry[1].encryptionData: {"checksumType":"SHA1/1K","checksum":"aa==","algorithm":{"name":"Blowfish CFB","initialisationVector":"bb=="},"keyDerivation":{"name":"PBKDF2","salt":"cc==","iterationCount":1024,"keySize":16}} != {"checksumType":"SHA1/1K","checksum":"aa==","algorithm":{"name":"Blowfish CFB","initialisationVector":"bb=="},"keyDerivation":{"name":"PBKDF2","salt":"cc==","iterationCount":2048,"keySize":16}}',
    ]);
  });

  it("distinguishes meaningful leading and trailing XML text whitespace", () => {
    const mimeType = "application/vnd.oasis.opendocument.text";
    const createPackage = (text: string) =>
      generateOcf(mimeType, { "content.xml": `<text:p>${text}</text:p>` });

    expect(compareOcfSemantic(createPackage("word "), createPackage("word"))).not.toEqual([]);
    expect(compareOcfSemantic(createPackage(" word"), createPackage("word"))).not.toEqual([]);
  });

  it("distinguishes CDATA content in XML parts", () => {
    const mimeType = "application/vnd.oasis.opendocument.text";
    const createPackage = (text: string) =>
      generateOcf(mimeType, { "content.xml": `<text:p><![CDATA[${text}]]></text:p>` });

    expect(compareOcfSemantic(createPackage("before"), createPackage("after"))).not.toEqual([]);
  });

  it("detects unmanifested package members in the semantic digest", () => {
    const mimeType = "application/vnd.oasis.opendocument.text";
    const source = generateOcf(mimeType, { "content.xml": "<text:p/>" });
    const changed = generateOcf(mimeType, {
      "content.xml": "<text:p/>",
      "unmanifested.bin": new Uint8Array([1, 2, 3]),
    });
    expect(compareOcfSemantic(source, changed)).toContainEqual(expect.stringContaining("member["));
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

  it("keeps manifest media types and omits missing declared paths", () => {
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

    const staleManifest = manifestXml(
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
    expect(staleManifest).not.toContain('manifest:full-path="Missing/data.bin"');
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
