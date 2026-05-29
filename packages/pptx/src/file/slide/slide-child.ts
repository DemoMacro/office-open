import type { ChartOptions } from "@file/chart/chart-frame";
import type { AudioFrameOptions } from "@file/media/audio-frame";
import type { VideoFrameOptions } from "@file/media/video-frame";
import type { PictureOptions } from "@file/picture/picture";
import type { GroupShapeOptions } from "@file/shape/group-shape";
import type { ConnectorShapeOptions, LineShapeOptions } from "@file/shape/line-shape";
import type { ShapeOptions } from "@file/shape/shape";
import type { SmartArtOptions } from "@file/smartart/smartart-frame";
import type { TableOptions } from "@file/table/table-frame";
import type { BaseXmlComponent } from "@file/xml-components";

/**
 * Discriminated union for slide/group children.
 * Accepts both class instances (backward compat) and plain objects (JSON-friendly).
 */
export type SlideChild =
  | BaseXmlComponent
  | { shape: ShapeOptions }
  | { picture: PictureOptions }
  | { table: TableOptions }
  | { chart: ChartOptions }
  | { line: LineShapeOptions }
  | { connector: ConnectorShapeOptions }
  | { video: VideoFrameOptions }
  | { audio: AudioFrameOptions }
  | { group: GroupShapeOptions }
  | { smartart: SmartArtOptions };
