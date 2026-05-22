import type { IChartFrameOptions } from "@file/chart/chart-frame";
import type { IAudioFrameOptions } from "@file/media/audio-frame";
import type { IVideoFrameOptions } from "@file/media/video-frame";
import type { IPictureOptions } from "@file/picture/picture";
import type { IGroupShapeOptions } from "@file/shape/group-shape";
import type { IConnectorShapeOptions, ILineShapeOptions } from "@file/shape/line-shape";
import type { IShapeOptions } from "@file/shape/shape";
import type { ISmartArtFrameOptions } from "@file/smartart/smartart-frame";
import type { ITableFrameOptions } from "@file/table/table-frame";
import type { BaseXmlComponent } from "@file/xml-components";

/**
 * Discriminated union for slide/group children.
 * Accepts both class instances (backward compat) and plain objects (JSON-friendly).
 */
export type SlideChild =
  | BaseXmlComponent
  | { shape: IShapeOptions }
  | { picture: IPictureOptions }
  | { table: ITableFrameOptions }
  | { chart: IChartFrameOptions }
  | { line: ILineShapeOptions }
  | { connector: IConnectorShapeOptions }
  | { video: IVideoFrameOptions }
  | { audio: IAudioFrameOptions }
  | { group: IGroupShapeOptions }
  | { smartart: ISmartArtFrameOptions };
