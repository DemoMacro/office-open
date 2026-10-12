import type { UniversalMeasure } from "@office-open/core";

export interface MediaTransformation {
  /** Display width in EMUs, or a universal measure such as `"25.4mm"`. */
  width: number | UniversalMeasure;
  /** Display height in EMUs, or a universal measure such as `"25.4mm"`. */
  height: number | UniversalMeasure;
}

interface CoreMediaData {
  fileName: string;
  transformation: MediaTransformation;
  data: Uint8Array;
}

interface RegularMediaData {
  type: "jpg" | "png" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf";
}

interface SvgMediaData {
  type: "svg";
  fallback: RegularMediaData & CoreMediaData;
}

interface VideoMediaData {
  type: "mp4" | "mov" | "wmv" | "avi";
}

interface AudioMediaData {
  type: "mp3" | "wav" | "wma" | "aac";
}

export type MediaData = (RegularMediaData | SvgMediaData | VideoMediaData | AudioMediaData) &
  CoreMediaData;
