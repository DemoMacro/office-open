import type { DataType } from "@office-open/core";

import type { MediaFrameBaseOptions } from "./media-frame-base";

export type AudioType = "mp3" | "wav" | "wma" | "aac";

/** A point on an audio CD (CT_AudioCDTime). */
export interface AudioCdTimeOptions {
  /** 1-based track number (required) */
  track: number;
  /** Offset within the track, in milliseconds (CT_AudioCDTime `@time`, default 0) */
  time?: number;
}

/** CD audio playback (a:audioCd) — the sound comes from an audio CD, not a media file. */
export interface AudioCdOptions {
  /** Playback start point (a:st, required) */
  start: AudioCdTimeOptions;
  /** Playback end point (a:end, required) */
  end: AudioCdTimeOptions;
}

export interface AudioFrameOptions extends Omit<MediaFrameBaseOptions, "data" | "type"> {
  /** Audio bytes — required for file-based audio, omitted for CD audio */
  data?: DataType;
  /** Audio format (required when data is set) */
  type?: AudioType;
  /** CD audio playback (a:audioCd) — mutually exclusive with data */
  audioCd?: AudioCdOptions;
  /** MIME content type of the linked audio (CT_AudioFile `@contentType`) */
  contentType?: string;
  /** Original audio file name (CT_EmbeddedWAVAudioFile `@name`, wav only) */
  audioFileName?: string;
}
