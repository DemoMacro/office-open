import type { ChartOptions } from "@parts/chart-frame";
import type { LockedCanvasFrameOptions } from "@parts/locked-canvas-frame";
import type { OleOptions } from "@parts/ole-frame";
import type { SmartArtOptions } from "@parts/smartart";
import type { AudioFrameOptions } from "@shared/media/audio-frame";
import type { VideoFrameOptions } from "@shared/media/video-frame";
import type { PictureOptions } from "@shared/picture";
import type { GroupShapeOptions } from "@shared/shape/group-shape";
import type { ConnectorShapeOptions, LineShapeOptions } from "@shared/shape/line-shape";
import type { ShapeOptions } from "@shared/shape/shape";
import type { TableOptions } from "@shared/table/table-frame";

/**
 * Discriminated union for slide/group children.
 * Each variant is a plain object identified by a unique key.
 */
export type SlideChild =
  | { shape: ShapeOptions }
  | { picture: PictureOptions }
  | { table: TableOptions }
  | { chart: ChartOptions }
  | { line: LineShapeOptions }
  | { connector: ConnectorShapeOptions }
  | { video: VideoFrameOptions }
  | { audio: AudioFrameOptions }
  | { group: GroupShapeOptions }
  | { smartart: SmartArtOptions }
  | { lockedCanvas: LockedCanvasFrameOptions }
  | { ole: OleOptions }
  | { contentPart: { rId: string } };
