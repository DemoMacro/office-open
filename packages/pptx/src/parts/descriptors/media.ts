/**
 * Video and Audio frame descriptors for PPTX.
 *
 * Both produce p:pic elements with media placeholders that are
 * resolved during compilation.
 *
 * @module
 */

import { convertToEmu } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import {
  stringifyNonVisualDrawingProperties,
  parseNonVisualDrawingProperties,
} from "@office-open/core/drawingml";
import { attr, attrNum, findChild, findFirst } from "@office-open/xml";
import type { AudioFrameOptions, AudioType } from "@shared/media/audio-frame";
import type { PosterType, VideoFrameOptions, VideoType } from "@shared/media/video-frame";

import { readPositionFromXfrm } from "./shape";

// ── ID counters ──

let _nextVideoId = 100;
let _nextAudioId = 200;

// ── Extension URIs ──

const VIDEO_EXT_URI = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";
const AUDIO_EXT_URI = "{CF1602FD-DB20-4165-A070-5F299619DA56}";

// ── Video descriptor ──

export const videoDesc: CustomDescriptor<VideoFrameOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const id = opts.id ?? _nextVideoId++;
    const name = opts.name ?? `Video ${id}`;
    const mediaFileName = `${name.replace(/\s+/g, "_")}.${opts.type ?? "mp4"}`;
    const posterFileName = `${name.replace(/\s+/g, "_")}_poster.${opts.posterType ?? "png"}`;

    const x = convertToEmu(opts.x ?? 0);
    const y = convertToEmu(opts.y ?? 0);
    const w = convertToEmu(opts.width ?? 0);
    const h = convertToEmu(opts.height ?? 0);

    const parts: string[] = [];

    // p:nvPicPr
    parts.push(
      `<p:nvPicPr>${stringifyNonVisualDrawingProperties("p:cNvPr", id, opts, name)}` +
        `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>` +
        `<p:nvPr><a:videoFile r:link="{video:${mediaFileName}}"/>` +
        `<p:extLst><p:ext uri="${VIDEO_EXT_URI}">` +
        `<p14:media r:embed="{media:${mediaFileName}}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"/>` +
        `</p:ext></p:extLst></p:nvPr></p:nvPicPr>`,
    );

    // p:blipFill (poster image)
    parts.push(
      `<p:blipFill><a:blip r:embed="{${posterFileName}}"/><a:stretch><a:fillRect/></a:stretch></p:blipFill>`,
    );

    // p:spPr
    parts.push(
      `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`,
    );

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
    const videoFileEl = findFirst(el, "a:videoFile");
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

  stringify(opts, _ctx) {
    const id = opts.id ?? _nextAudioId++;
    const name = opts.name ?? `Audio ${id}`;
    const mediaFileName = `${name.replace(/\s+/g, "_")}.${opts.type ?? "mp3"}`;

    const x = convertToEmu(opts.x ?? 0);
    const y = convertToEmu(opts.y ?? 0);
    const w = convertToEmu(opts.width ?? 0);
    const h = convertToEmu(opts.height ?? 0);

    const parts: string[] = [];

    // p:nvPicPr
    parts.push(
      `<p:nvPicPr>${stringifyNonVisualDrawingProperties("p:cNvPr", id, opts, name)}` +
        `<p:cNvPicPr><a:picLocks noChangeAspect="1"/></p:cNvPicPr>` +
        `<p:nvPr><p:extLst><p:ext uri="${AUDIO_EXT_URI}">` +
        `<p14:media r:embed="{media:${mediaFileName}}" xmlns:p14="http://schemas.microsoft.com/office/powerpoint/2010/main"/>` +
        `</p:ext></p:extLst></p:nvPr></p:nvPicPr>`,
    );

    // p:blipFill (no poster for audio)
    parts.push(`<p:blipFill><a:stretch><a:fillRect/></a:stretch></p:blipFill>`);

    // p:spPr
    parts.push(
      `<p:spPr><a:xfrm><a:off x="${x}" y="${y}"/><a:ext cx="${w}" cy="${h}"/></a:xfrm>` +
        `<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></p:spPr>`,
    );

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

    // Media data from a:audioFile (r:link) or p14:media (r:embed)
    const audioFileEl = findFirst(el, "a:audioFile");
    const rLink = audioFileEl ? attr(audioFileEl, "r:link") : undefined;
    const p14media = !audioFileEl ? findFirst(el, "p14:media") : undefined;
    const rEmbed = p14media ? attr(p14media, "r:embed") : undefined;
    const mediaRef = rLink ?? rEmbed;
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

    return result as AudioFrameOptions;
  },
};

// ── Helpers ──

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
