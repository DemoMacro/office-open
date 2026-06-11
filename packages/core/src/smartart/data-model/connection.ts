import { escapeXml } from "@office-open/xml";

/**
 * dgm:cxn — SmartArt data model connection (edge) XML stringifier.
 */
export function stringifyConnection(opts: {
  modelId: string | number;
  srcId: string | number;
  destId: string | number;
  type?: string;
  srcOrd?: number;
  destOrd?: number;
  parTransId?: string;
  sibTransId?: string;
  presId?: string;
}): string {
  const attrs: string[] = [
    `modelId="${escapeXml(String(opts.modelId))}"`,
    `srcId="${escapeXml(String(opts.srcId))}"`,
    `destId="${escapeXml(String(opts.destId))}"`,
    `srcOrd="${opts.srcOrd ?? 0}"`,
    `destOrd="${opts.destOrd ?? 0}"`,
  ];
  if (opts.type) attrs.push(`type="${escapeXml(opts.type)}"`);
  if (opts.parTransId) attrs.push(`parTransId="${escapeXml(opts.parTransId)}"`);
  if (opts.sibTransId) attrs.push(`sibTransId="${escapeXml(opts.sibTransId)}"`);
  if (opts.presId) attrs.push(`presId="${escapeXml(opts.presId)}"`);
  return `<dgm:cxn ${attrs.join(" ")}/>`;
}
