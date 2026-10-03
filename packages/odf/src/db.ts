import { ODF_NAMESPACES, metaXml, parseMeta } from "./meta";
import { parseOdfNodes, serializeOdfNodes, type OdfXmlNode } from "./odf-node";
import { generateOcf, readOcf, readXml } from "./package";
import { childNamed } from "./xml";

const MIME = "application/vnd.oasis.opendocument.database";
const NAMESPACES = `${ODF_NAMESPACES} xmlns:db="urn:oasis:names:tc:opendocument:xmlns:database:1.0"`;

export interface DatabaseDocumentOptions {
  title?: string;
  body: OdfXmlNode[];
}

export function generateDatabaseDocument(options: DatabaseDocumentOptions): Uint8Array {
  const body = serializeOdfNodes(options.body).join("");
  return generateOcf(MIME, {
    "content.xml": databaseContentXml(body),
    "meta.xml": metaXml({ title: options.title }),
  });
}

export function parseDatabaseDocument(data: Uint8Array): DatabaseDocumentOptions {
  const { files } = readOcf(data, MIME);
  const content = readXml(files, "content.xml");
  const body = childNamed(childNamed(content, "office:body"), "office:database");
  return { ...parseMeta(files), body: parseOdfNodes(body) };
}

function databaseContentXml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.3"><office:body><office:database>${body}</office:database></office:body></office:document-content>`;
}
