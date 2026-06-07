export interface PictureOptions {
  readonly x?: number;
  readonly y?: number;
  readonly width?: number;
  readonly height?: number;
  readonly data: Uint8Array;
  readonly type: "png" | "jpg" | "gif" | "bmp" | "emf" | "wmf";
  readonly name?: string;
}
