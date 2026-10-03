import { ODF_NAMESPACES, metaXml, parseMeta } from "./meta";
import { parseOdfNodes, serializeOdfNodes, type OdfXmlNode } from "./odf-node";
import { generateOcf, readOcf, readXml } from "./package";
import { childNamed } from "./xml";

const MIME = "application/vnd.oasis.opendocument.chart";
const NAMESPACES = `${ODF_NAMESPACES} xmlns:chart="urn:oasis:names:tc:opendocument:xmlns:chart:1.0"`;

export interface ChartDocumentOptions {
  title?: string;
  body: OdfXmlNode[];
}

export function generateChartDocument(options: ChartDocumentOptions): Uint8Array {
  const body = serializeOdfNodes(options.body).join("");
  return generateOcf(MIME, {
    "content.xml": chartContentXml(body),
    "meta.xml": metaXml({ title: options.title }),
  });
}

export function parseChartDocument(data: Uint8Array): ChartDocumentOptions {
  const { files } = readOcf(data, MIME);
  const content = readXml(files, "content.xml");
  const body = childNamed(childNamed(content, "office:body"), "office:chart");
  return { ...parseMeta(files), body: parseOdfNodes(body) };
}

function chartContentXml(body: string): string {
  return `<?xml version="1.0" encoding="UTF-8"?><office:document-content ${NAMESPACES} office:version="1.3"><office:body><office:chart>${body}</office:chart></office:body></office:document-content>`;
}
