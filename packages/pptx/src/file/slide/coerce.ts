import type { MasterChild } from "@file/file";
import { Shape } from "@file/shape/shape";
import type { Context } from "@file/xml-components";
import { BaseXmlComponent } from "@file/xml-components";

function coerceMasterChild(child: MasterChild): BaseXmlComponent {
  if (child instanceof BaseXmlComponent) return child;
  if ("shape" in child) return new Shape(child.shape);
  throw new Error("Unknown master child type");
}

export function buildMasterChildrenXml(children?: readonly MasterChild[]): string {
  if (!children || children.length === 0) return "";
  const ctx: Context = { stack: [] };
  let result = "";
  for (const child of children) {
    const xmlStr = coerceMasterChild(child).toXml(ctx);
    if (xmlStr) result += xmlStr;
  }
  return result;
}
