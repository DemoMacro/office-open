import { ChartFrame } from "@file/chart/chart-frame";
import type { MasterChild } from "@file/file";
import { AudioFrame } from "@file/media/audio-frame";
import { VideoFrame } from "@file/media/video-frame";
import { Picture } from "@file/picture/picture";
import { GroupShape } from "@file/shape/group-shape";
import { ConnectorShape, LineShape } from "@file/shape/line-shape";
import { Shape } from "@file/shape/shape";
import type { SlideChild } from "@file/slide/slide-child";
import { SmartArtFrame } from "@file/smartart/smartart-frame";
import { TableFrame } from "@file/table/table-frame";
import { BaseXmlComponent } from "@file/xml-components";

export function coerceChild(child: SlideChild): BaseXmlComponent {
  if (child instanceof BaseXmlComponent) return child;
  if ("shape" in child) return new Shape(child.shape);
  if ("picture" in child) return new Picture(child.picture);
  if ("table" in child) return new TableFrame(child.table);
  if ("chart" in child) return new ChartFrame(child.chart);
  if ("line" in child) return new LineShape(child.line);
  if ("connector" in child) return new ConnectorShape(child.connector);
  if ("video" in child) return new VideoFrame(child.video);
  if ("audio" in child) return new AudioFrame(child.audio);
  if ("group" in child) return new GroupShape(child.group);
  if ("smartart" in child) return new SmartArtFrame(child.smartart);
  throw new Error("Unknown slide child type");
}

export function coerceMasterChild(child: MasterChild): BaseXmlComponent {
  if (child instanceof BaseXmlComponent) return child;
  if ("shape" in child) return new Shape(child.shape);
  throw new Error("Unknown master child type");
}
