export interface PictureOptions {
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  data: Uint8Array;
  type: "png" | "jpg" | "gif" | "bmp" | "emf" | "wmf";
  name?: string;
}
