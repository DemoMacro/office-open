import {
  generateOcf,
  manifestOptionsXml,
  parseManifestOptions,
} from "../../../packages/odf/dist/index.mjs";
import { parseDocument as parseOdtDocument } from "../../../packages/odt/dist/index.mjs";
import { ENCODER, assert, assertEqual } from "./support";

export async function odtMissingEmbeddedObject(): Promise<void> {
  const part = "content.xml draw:object";
  const namespaces = `office="urn:oasis:names:tc:opendocument:xmlns:office:1.0" draw="urn:oasis:names:tc:opendocument:xmlns:drawing:1.0" text="urn:oasis:names:tc:opendocument:xmlns:text:1.0" xlink="http://www.w3.org/1999/xlink" svg="urn:oasis:names:tc:opendocument:xmlns:svg-compatible:1.0" style="urn:oasis:names:tc:opendocument:xmlns:style:1.0"`;
  const source = generateOcf(
    "application/vnd.oasis.opendocument.text",
    {
      "content.xml": ENCODER.encode(
        `<?xml version="1.0"?><office:document-content ${namespaces}><office:automatic-styles/><office:body><office:text><text:p><draw:frame svg:width="4cm" svg:height="3cm"><draw:object xlink:href="./Object 1" xlink:type="simple" xlink:show="embed" xlink:actuate="onLoad"/></draw:frame></text:p></office:text></office:body></office:document-content>`,
      ),
      "styles.xml": ENCODER.encode(
        `<?xml version="1.0"?><office:document-styles ${namespaces}><office:styles/></office:document-styles>`,
      ),
      "Object 1/": new Uint8Array(),
    },
    {},
    {
      version: "1.3",
      entries: [
        { fullPath: "/", version: "1.3", mediaType: "application/vnd.oasis.opendocument.text" },
        { fullPath: "content.xml", mediaType: "text/xml" },
        { fullPath: "styles.xml", mediaType: "text/xml" },
        { fullPath: "Object 1/", mediaType: "application/vnd.oasis.opendocument.presentation" },
      ],
    },
  );
  let error: unknown;
  try {
    parseOdtDocument(source);
  } catch (cause) {
    error = cause;
  }
  const diagnostic = error as { name?: string; reason?: string } | undefined;
  assert(diagnostic, part, "missing embedded object did not reject");
  assertEqual(diagnostic?.name, "draw:object", part, "embedded object error element");
  assertEqual(
    diagnostic?.reason,
    "embedded object subdocument is missing",
    part,
    "embedded object error reason",
  );
}

const MANIFEST_NAMESPACE = `xmlns:manifest="urn:oasis:names:tc:opendocument:xmlns:manifest:1.0"`;

const MINIMAL_MANIFEST = `<?xml version="1.0"?><manifest:manifest ${MANIFEST_NAMESPACE}><manifest:file-entry manifest:full-path="/" manifest:media-type="application/vnd.oasis.opendocument.text"/></manifest:manifest>`;

const FULL_MANIFEST = `<?xml version="1.0"?><manifest:manifest ${MANIFEST_NAMESPACE} manifest:version="1.3"><manifest:encrypted-key><manifest:encryption-method manifest:PGPAlgorithm="RSA"/><manifest:keyinfo><manifest:PGPData><manifest:PGPKeyID>MQ==</manifest:PGPKeyID><manifest:PGPKeyPacket>Ag==</manifest:PGPKeyPacket></manifest:PGPData></manifest:keyinfo><manifest:CipherData><manifest:CipherValue>Aw==</manifest:CipherValue></manifest:CipherData></manifest:encrypted-key><manifest:file-entry manifest:full-path="/" manifest:version="1.3" manifest:media-type="application/vnd.oasis.opendocument.text"/><manifest:file-entry manifest:full-path="content.xml" manifest:media-type="text/xml" manifest:preferred-view-mode="page" manifest:size="7"><manifest:encryption-data manifest:checksum-type="SHA1/1K" manifest:checksum="aa=="><manifest:algorithm manifest:algorithm-name="Blowfish CFB" manifest:initialisation-vector="bb=="/><manifest:key-derivation manifest:key-derivation-name="PBKDF2" manifest:salt="cc==" manifest:iteration-count="1024" manifest:key-size="16"/><manifest:start-key-generation manifest:start-key-generation-name="SHA256" manifest:key-size="32"/></manifest:encryption-data></manifest:file-entry></manifest:manifest>`;

export async function odfManifestMatrix(): Promise<void> {
  const part = "META-INF/manifest.xml";
  for (const source of [MINIMAL_MANIFEST, FULL_MANIFEST]) {
    const parsed = parseManifestOptions(source);
    assertEqual(
      JSON.stringify(parseManifestOptions(manifestOptionsXml(parsed))),
      JSON.stringify(parsed),
      part,
      "manifest property matrix round-trip",
    );
  }
  const parsed = parseManifestOptions(FULL_MANIFEST);
  assertEqual(parsed.version, "1.3", part, "manifest version");
  assertEqual(parsed.entries[1]?.preferredViewMode, "page", part, "preferred view mode");
  assertEqual(parsed.entries[1]?.size, 7, part, "entry size");
  assertEqual(
    parsed.entries[1]?.encryptionData?.startKeyGeneration?.keySize,
    32,
    part,
    "start key size",
  );
  assertEqual(parsed.encryptedKeys?.[0]?.algorithm, "RSA", part, "encrypted key algorithm");
}
