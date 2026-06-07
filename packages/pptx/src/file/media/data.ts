export interface MediaDataTransformation {
  pixels: {
    x: number;
    y: number;
  };
  emus: {
    x: number;
    y: number;
  };
  flip?: {
    vertical?: boolean;
    horizontal?: boolean;
  };
  rotation?: number;
}

interface CoreMediaData {
  fileName: string;
  transformation: MediaDataTransformation;
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

export type IMediaData = (RegularMediaData | SvgMediaData | VideoMediaData | AudioMediaData) &
  CoreMediaData;
