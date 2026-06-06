import { Chart } from "@file/chart/chart-frame";
import type { MasterChild } from "@file/file";
import { LockedCanvasFrame } from "@file/locked-canvas/locked-canvas-frame";
import { AudioFrame } from "@file/media/audio-frame";
import { VideoFrame } from "@file/media/video-frame";
import { OleFrame } from "@file/ole/ole-frame";
import { Picture } from "@file/picture/picture";
import { GroupShape } from "@file/shape/group-shape";
import { ConnectorShape, LineShape } from "@file/shape/line-shape";
import { Shape } from "@file/shape/shape";
import type { SlideChild } from "@file/slide/slide-child";
import { SmartArt } from "@file/smartart/smartart-frame";
import { Table } from "@file/table/table-frame";
import type { Context } from "@file/xml-components";
import { BaseXmlComponent } from "@file/xml-components";

class ContentPart extends BaseXmlComponent {
  private readonly rId: string;

  public constructor(rId: string) {
    super("p:contentPart");
    this.rId = rId;
  }

  public override toXml(_context: Context): string {
    return `<p:contentPart r:id="${this.rId}"/>`;
  }
}

export function coerceChild(child: SlideChild): BaseXmlComponent {
  if (child instanceof BaseXmlComponent) return child;
  if ("shape" in child) return new Shape(child.shape);
  if ("picture" in child) return new Picture(child.picture);
  if ("table" in child) return new Table(child.table);
  if ("chart" in child) return new Chart(child.chart);
  if ("line" in child) return new LineShape(child.line);
  if ("connector" in child) return new ConnectorShape(child.connector);
  if ("video" in child) return new VideoFrame(child.video);
  if ("audio" in child) return new AudioFrame(child.audio);
  if ("group" in child) return new GroupShape(child.group);
  if ("smartart" in child) return new SmartArt(child.smartart);
  if ("lockedCanvas" in child) return new LockedCanvasFrame(child.lockedCanvas);
  if ("ole" in child) return new OleFrame(child.ole);
  if ("contentPart" in child) {
    return new ContentPart(child.contentPart.rId);
  }
  throw new Error("Unknown slide child type");
}

export function coerceMasterChild(child: MasterChild): BaseXmlComponent {
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
