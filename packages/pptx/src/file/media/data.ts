export interface IMediaDataTransformation {
    readonly pixels: {
        readonly x: number;
        readonly y: number;
    };
    readonly emus: {
        readonly x: number;
        readonly y: number;
    };
    readonly flip?: {
        readonly vertical?: boolean;
        readonly horizontal?: boolean;
    };
    readonly rotation?: number;
}

interface CoreMediaData {
    readonly fileName: string;
    readonly transformation: IMediaDataTransformation;
    readonly data: Uint8Array;
}

interface RegularMediaData {
    readonly type: "jpg" | "png" | "gif" | "bmp" | "tif" | "ico" | "emf" | "wmf";
}

interface SvgMediaData {
    readonly type: "svg";
    readonly fallback: RegularMediaData & CoreMediaData;
}

export type IMediaData = (RegularMediaData | SvgMediaData) & CoreMediaData;
