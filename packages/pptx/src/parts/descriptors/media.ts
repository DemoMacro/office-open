/**
 * Video and Audio frame descriptors for PPTX.
 *
 * Both produce p:pic elements with media placeholders that are
 * resolved during compilation.
 *
 * Reference: OOXML transitional, dml-main.xsd EG_Media
 * (audioCd / wavAudioFile / audioFile / videoFile / quickTimeFile).
 *
 * @module
 */

import { toUint8Array } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import {
  shapePropertiesDesc,
  stringifyNonVisualDrawingProperties,
  parseNonVisualDrawingProperties,
} from "@office-open/core/drawingml";
import { attr, attrNum, escapeXml, findChild, findFirst } from "@office-open/xml";
import type { Element } from "@office-open/xml";
import type { AudioCdOptions, AudioFrameOptions, AudioType } from "@shared/media/audio-frame";
import type { MediaFrameBaseOptions } from "@shared/media/media-frame-base";
import type { PosterType, VideoFrameOptions, VideoType } from "@shared/media/video-frame";

import type { MediaEntry, PptxWriteContext } from "../../context";
import { readPositionFromXfrm } from "./shape";

// ── ID counters ──

let _nextVideoId = 100;
let _nextAudioId = 200;

// ── Extension URIs ──

const VIDEO_EXT_URI = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";
const AUDIO_EXT_URI = "{CF1602FD-DB20-4165-A070-5F2996DA56}";

// ── Shared media element builders (EG_Media) ──

/** a:audioCd — CD track playback, no media file. */
function stringifyAudioCd(cd: AudioCdOptions): string {
  return (
    `<a:audioCd><a:st${stringifyAudioCdTime(cd.start)}/>` +
    `<a:end${stringifyAudioCdTime(cd.end)}/></a:audioCd>`
  );
}

function stringifyAudioCdTime(t: AudioCdOptions["start"]): string {
  let s = ` track="${t.track}"`;
  if (t.time !== undefined) s += ` time="${t.time}"`;
  return s;
}

/** a:audioFile / a:wavAudioFile — linked/embedded audio file. */
function stringifyAudioFile(
  mediaFileName: string,
  type: AudioType,
  opts: AudioFrameOptions,
): string {
  const contentType = opts.contentType ? ` contentType="${escapeXml(opts.contentType)}"` : "";
  if (type === "wav") {
    // Embedded WAV: r:embed + optional original file name.
    const name = opts.audioFileName ? ` name="${escapeXml(opts.audioFileName)}"` : "";
    return `<a:wavAudioFile r:embed="{audio:${mediaFileName}}"${name}/>`;
  }
  return `<a:audioFile r:link="{audio:${mediaFileName}}"${contentType}/>`;
}

// ── Video descriptor ──

export const videoDesc: CustomDescriptor<VideoFrameOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const id = opts.id ?? _nextVideoId++;
    const name = opts.name ?? `Video ${id}`;
    const pptx = ctx as PptxWriteContext;
    const mediaFileName = registerMediaFile(
      pptx,
      opts.data,
      opts.type ?? "mp4",
      `${name.replace(/\s+/g, "_")}.${opts.type ?? "mp4"}`,
    );

    const parts: string[] = [];

    // p:nvPicPr
    const contentType = opts.contentType ? ` contentType="${escapeXml(opts.contentType)}"` : "";
    const mediaEl = mediaFileName
      ? `<a:videoFile r:link="{video:${mediaFileName}}"${contentType}/>`
      : "";
    parts.push(
      `<p:nvPicPr>${stringifyNonVisualDrawingProperties("p:cNvPr", id, opts, name)}` +
        `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>` +
        `<p:nvPr>${mediaEl}` +
        (mediaFileName
          ? `<p:extLst><p:ext uri="${VIDEO_EXT_URI}">` +
            `<p14:media r:embed="{media:${mediaFileName}}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"/>` +
            `</p:ext></p:extLst>`
          : "") +
        `</p:nvPr></p:nvPicPr>`,
    );

    // p:blipFill (poster image)
    let posterAttr = "";
    if (opts.poster) {
      const posterFileName = registerMediaFile(
        pptx,
        opts.poster,
        opts.posterType ?? "png",
        `${name.replace(/\s+/g, "_")}_poster.${opts.posterType ?? "png"}`,
      );
      if (posterFileName) posterAttr = `<a:blip r:embed="{${posterFileName}}"/>`;
    }
    parts.push(`<p:blipFill>${posterAttr}<a:stretch><a:fillRect/></a:stretch></p:blipFill>`);

    // p:spPr
    parts.push(stringifyMediaSpPr(opts, ctx));

    return `<p:pic>${parts.join("")}</p:pic>`;
  },

  parse(el, _ctx) {
    const result: Partial<VideoFrameOptions> = {};

    // Position from p:spPr
    const spPr = findChild(el, "p:spPr");
    if (spPr) {
      const xfrm = findChild(spPr, "a:xfrm");
      if (xfrm) Object.assign(result, readPositionFromXfrm(xfrm));
    }

    // id + name from p:nvPicPr → a:cNvPr or p:cNvPr
    const nvPicPr = findChild(el, "p:nvPicPr");
    if (nvPicPr) {
      const cNvPr = findChild(nvPicPr, "a:cNvPr") ?? findChild(nvPicPr, "p:cNvPr");
      if (cNvPr) {
        Object.assign(result, parseNonVisualDrawingProperties(cNvPr));
        const id = attrNum(cNvPr, "id");
        if (id !== undefined) result.id = id;
      }
    }

    // Media data from a:videoFile (r:link) or p14:media (r:embed)
    const quickTimeEl = findFirst(el, "a:quickTimeFile");
    const videoFileEl = findFirst(el, "a:videoFile") ?? quickTimeEl;
    const rLink = videoFileEl ? attr(videoFileEl, "r:link") : undefined;
    const p14media = !videoFileEl ? findFirst(el, "p14:media") : undefined;
    const rEmbed = p14media ? attr(p14media, "r:embed") : undefined;
    const mediaRef = rLink ?? rEmbed;
    if (mediaRef) {
      const mediaPath = _ctx.resolveRelationship(mediaRef);
      if (mediaPath) {
        const data = _ctx.getRaw(mediaPath);
        if (data) result.data = data;
      }
      // Fall back to the placeholder reference (e.g. "{media:foo.mp4}") when the
      // relationship isn't registered, so the type survives a round-trip.
      result.type = mediaTypeFromPath(mediaPath ?? mediaRef, "video");
    }
    if (quickTimeEl) result.type = "mov";
    if (videoFileEl) {
      const contentType = attr(videoFileEl, "contentType");
      if (contentType) result.contentType = contentType;
    }

    // Poster from blipFill
    const blipFill = findChild(el, "p:blipFill");
    if (blipFill) {
      const blip = findChild(blipFill, "a:blip");
      if (blip) {
        const rEmbedPoster = attr(blip, "r:embed");
        if (rEmbedPoster) {
          const posterPath = _ctx.resolveRelationship(rEmbedPoster);
          if (posterPath) {
            const posterData = _ctx.getRaw(posterPath);
            if (posterData) result.poster = posterData;
            result.posterType = imageTypeFromPath(posterPath) as PosterType;
          }
        }
      }
    }

    return result as VideoFrameOptions;
  },
};

// ── Audio descriptor ──

export const audioDesc: CustomDescriptor<AudioFrameOptions> = {
  kind: "custom",

  stringify(opts, ctx) {
    const id = opts.id ?? _nextAudioId++;
    const name = opts.name ?? `Audio ${id}`;
    const pptx = ctx as PptxWriteContext;
    const mediaFileName =
      opts.data !== undefined
        ? registerMediaFile(
            pptx,
            opts.data,
            opts.type ?? "mp3",
            `${name.replace(/\s+/g, "_")}.${opts.type ?? "mp3"}`,
          )
        : undefined;

    const parts: string[] = [];

    // p:nvPicPr — EG_Media choice, then the p14:media extension
    let mediaEl: string;
    if (opts.cd) {
      mediaEl = stringifyAudioCd(opts.cd);
    } else if (mediaFileName) {
      mediaEl = stringifyAudioFile(mediaFileName, opts.type ?? "mp3", opts);
    } else {
      mediaEl = "";
    }
    parts.push(
      `<p:nvPicPr>${stringifyNonVisualDrawingProperties("p:cNvPr", id, opts, name)}` +
        `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>` +
        `<p:nvPr>${mediaEl}` +
        (mediaFileName
          ? `<p:extLst><p:ext uri="${AUDIO_EXT_URI}">` +
            `<p14:media r:embed="{media:${mediaFileName}}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"/>` +
            `</p:ext></p:extLst>`
          : "") +
        `</p:nvPr></p:nvPicPr>`,
    );

    // p:blipFill (no poster for audio)
    parts.push(`<p:blipFill><a:stretch><a:fillRect/></a:stretch></p:blipFill>`);

    // p:spPr
    parts.push(stringifyMediaSpPr(opts, ctx));

    return `<p:pic>${parts.join("")}</p:pic>`;
  },

  parse(el, _ctx) {
    const result: Partial<AudioFrameOptions> = {};

    // Position from p:spPr
    const spPr = findChild(el, "p:spPr");
    if (spPr) {
      const xfrm = findChild(spPr, "a:xfrm");
      if (xfrm) Object.assign(result, readPositionFromXfrm(xfrm));
    }

    // id + name from p:nvPicPr
    const nvPicPr = findChild(el, "p:nvPicPr");
    if (nvPicPr) {
      const cNvPr = findChild(nvPicPr, "a:cNvPr") ?? findChild(nvPicPr, "p:cNvPr");
      if (cNvPr) {
        Object.assign(result, parseNonVisualDrawingProperties(cNvPr));
        const id = attrNum(cNvPr, "id");
        if (id !== undefined) result.id = id;
      }
    }

    // CD audio (a:audioCd) — track/time, no media file
    const audioCdEl = findFirst(el, "a:audioCd");
    if (audioCdEl) {
      const stEl = findChild(audioCdEl, "a:st");
      const endEl = findChild(audioCdEl, "a:end");
      if (stEl && endEl) {
        const cd: AudioCdOptions = {
          start: readAudioCdTime(stEl),
          end: readAudioCdTime(endEl),
        };
        result.cd = cd;
      }
      return result as AudioFrameOptions;
    }

    // Media data from a:audioFile / a:wavAudioFile (r:link/r:embed) or p14:media
    const wavFileEl = findFirst(el, "a:wavAudioFile");
    const audioFileEl = findFirst(el, "a:audioFile") ?? wavFileEl;
    const rLink = audioFileEl ? attr(audioFileEl, "r:link") : undefined;
    const rEmbedAttr = audioFileEl ? attr(audioFileEl, "r:embed") : undefined;
    const p14media = !audioFileEl ? findFirst(el, "p14:media") : undefined;
    const rEmbedExt = p14media ? attr(p14media, "r:embed") : undefined;
    const mediaRef = rLink ?? rEmbedAttr ?? rEmbedExt;
    if (mediaRef) {
      const mediaPath = _ctx.resolveRelationship(mediaRef);
      if (mediaPath) {
        const data = _ctx.getRaw(mediaPath);
        if (data) result.data = data;
      }
      // Fall back to the placeholder reference (e.g. "{media:foo.wav}") when the
      // relationship isn't registered, so the type survives a round-trip.
      result.type = mediaTypeFromPath(mediaPath ?? mediaRef, "audio");
    }
    if (wavFileEl) result.type = "wav";
    if (audioFileEl) {
      const contentType = attr(audioFileEl, "contentType");
      if (contentType) result.contentType = contentType;
      const audioFileName = attr(audioFileEl, "name");
      if (audioFileName) result.audioFileName = audioFileName;
    }

    return result as AudioFrameOptions;
  },
};

// ── Helpers ──

function readAudioCdTime(el: Element): AudioCdOptions["start"] {
  const time = attrNum(el, "time");
  return { track: attrNum(el, "track") ?? 0, ...(time !== undefined ? { time } : {}) };
}

/**
 * Register media bytes under a deterministic file name and return the canonical
 * name (which may differ when the requested name is taken by other bytes).
 * Returns undefined when data is undefined.
 */
function registerMediaFile(
  pptx: PptxWriteContext,
  data: MediaFrameBaseOptions["data"] | undefined,
  type: string,
  fileName: string,
): string | undefined {
  if (data === undefined) return undefined;
  const raw = toUint8Array(data, { encoding: "base64" });
  const entry: MediaEntry = {
    key: fileName,
    data: raw,
    fileName,
    type,
    transformation: { pixels: { x: 0, y: 0 }, emus: { x: 0, y: 0 } },
  };
  return pptx.addImage(fileName, entry).fileName;
}

/** p:spPr for a media frame: position/size + rect geometry (picture precedent). */
function stringifyMediaSpPr(
  opts: Pick<MediaFrameBaseOptions, "x" | "y" | "width" | "height">,
  ctx: Parameters<typeof shapePropertiesDesc.stringify>[1],
): string {
  const content = shapePropertiesDesc.stringify(
    {
      x: opts.x ?? 0,
      y: opts.y ?? 0,
      width: opts.width ?? 0,
      height: opts.height ?? 0,
      geometry: "rect",
    },
    ctx,
  );
  return `<p:spPr>${content ?? ""}</p:spPr>`;
}

function mediaTypeFromPath(path: string, kind: "video"): VideoType;
function mediaTypeFromPath(path: string, kind: "audio"): AudioType;
function mediaTypeFromPath(path: string, kind: "video" | "audio"): VideoType | AudioType {
  // Tolerate placeholder wrappers like "{media:foo.wav}" — extract the last
  // alphanumeric extension token rather than a trailing fragment with "}".
  const ext = path.match(/\.([a-z0-9]+)\b/i)?.[1]?.toLowerCase() ?? "";
  if (kind === "video") {
    if (["mp4", "mov", "wmv", "avi"].includes(ext)) return ext as VideoType;
    return "mp4";
  }
  if (["mp3", "wav", "wma", "aac"].includes(ext)) return ext as AudioType;
  return "mp3";
}

function imageTypeFromPath(path: string): string {
  const ext = path.split(".").pop()?.toLowerCase() ?? "";
  if (["png", "jpg", "jpeg", "gif", "bmp", "svg"].includes(ext)) return ext;
  return "png";
}
