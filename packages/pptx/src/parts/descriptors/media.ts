/**
 * Video and Audio frame descriptors for PPTX.
 *
 * Both produce p:pic elements with media placeholders that are
 * resolved during compilation.
 *
 * @module
 */

import { convertPixelsToEmu } from "@office-open/core";
import type { DataType } from "@office-open/core";
import type { CustomDescriptor } from "@office-open/core/descriptor";
import { attr, findChild, findDeep } from "@office-open/xml";
import { escapeXml } from "@office-open/xml";

import { readPositionFromXfrm } from "./shape";

// ── Types ──

export type VideoType = "mp4" | "mov" | "wmv" | "avi";
export type AudioType = "mp3" | "wav" | "wma" | "aac";
export type PosterType = "png" | "jpg";

export interface VideoDescriptorOptions {
  id?: number;
  name?: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  data?: DataType;
  type?: VideoType;
  poster?: DataType;
  posterType?: PosterType;
}

export interface AudioDescriptorOptions {
  id?: number;
  name?: string;
  x?: number;
  y?: number;
  width?: number;
  height?: number;
  data?: DataType;
  type?: AudioType;
}

// ── ID counters ──

let _nextVideoId = 100;
let _nextAudioId = 200;

// ── Extension URIs ──

const VIDEO_EXT_URI = "{DAA4B4D4-6D71-4841-9C94-3DE7FCFB9230}";
const AUDIO_EXT_URI = "{CF1602FD-DB20-4165-A070-5F299619DA56}";

// ── Video descriptor ──

export const videoDesc: CustomDescriptor<VideoDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const id = opts.id ?? _nextVideoId++;
    const name = opts.name ?? `Video ${id}`;
    const mediaFileName = `${name.replace(/\s+/g, "_")}.${opts.type ?? "mp4"}`;
    const posterFileName = `${name.replace(/\s+/g, "_")}_poster.${opts.posterType ?? "png"}`;

    const x = convertPixelsToEmu(opts.x ?? 0);
    const y = convertPixelsToEmu(opts.y ?? 0);
    const w = convertPixelsToEmu(opts.width ?? 0);
    const h = convertPixelsToEmu(opts.height ?? 0);

    const parts: string[] = [];

    // p:nvPicPr
    parts.push(
      `<p:nvPicPr><p:cNvPr id="${id}" name="${escapeXml(name)}" descr=""/>` +
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
    const result: Partial<VideoDescriptorOptions> = {};

    // Position from p:spPr
    const spPr = findChild(el, "p:spPr");
    if (spPr) {
      const xfrm = findChild(spPr, "a:xfrm");
      if (xfrm) Object.assign(result, readPositionFromXfrm(xfrm));
    }

    // Name from p:nvPicPr → a:cNvPr or p:cNvPr
    const nvPicPr = findChild(el, "p:nvPicPr");
    if (nvPicPr) {
      const cNvPr = findChild(nvPicPr, "a:cNvPr") ?? findChild(nvPicPr, "p:cNvPr");
      if (cNvPr) {
        const name = attr(cNvPr, "name");
        if (name) result.name = name;
      }
    }

    // Media data from a:videoFile (r:link) or p14:media (r:embed)
    const videoFileEl = findDeep(el, "a:videoFile")[0];
    const rLink = videoFileEl ? attr(videoFileEl, "r:link") : undefined;
    const p14media = !videoFileEl ? findDeep(el, "p14:media")[0] : undefined;
    const rEmbed = p14media ? attr(p14media, "r:embed") : undefined;
    const mediaRef = rLink ?? rEmbed;
    if (mediaRef) {
      const mediaPath = _ctx.resolveRelationship(mediaRef);
      if (mediaPath) {
        const data = _ctx.getRaw(mediaPath);
        if (data) result.data = data;
        result.type = mediaTypeFromPath(mediaPath, "video");
      }
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

    return result as VideoDescriptorOptions;
  },
};

// ── Audio descriptor ──

export const audioDesc: CustomDescriptor<AudioDescriptorOptions> = {
  kind: "custom",

  stringify(opts, _ctx) {
    const id = opts.id ?? _nextAudioId++;
    const name = opts.name ?? `Audio ${id}`;
    const mediaFileName = `${name.replace(/\s+/g, "_")}.${opts.type ?? "mp3"}`;

    const x = convertPixelsToEmu(opts.x ?? 0);
    const y = convertPixelsToEmu(opts.y ?? 0);
    const w = convertPixelsToEmu(opts.width ?? 0);
    const h = convertPixelsToEmu(opts.height ?? 0);

    const parts: string[] = [];

    // p:nvPicPr
    parts.push(
      `<p:nvPicPr><p:cNvPr id="${id}" name="${escapeXml(name)}" descr=""/>` +
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
    const result: Partial<AudioDescriptorOptions> = {};

    // Position from p:spPr
    const spPr = findChild(el, "p:spPr");
    if (spPr) {
      const xfrm = findChild(spPr, "a:xfrm");
      if (xfrm) Object.assign(result, readPositionFromXfrm(xfrm));
    }

    // Name from p:nvPicPr
    const nvPicPr = findChild(el, "p:nvPicPr");
    if (nvPicPr) {
      const cNvPr = findChild(nvPicPr, "a:cNvPr") ?? findChild(nvPicPr, "p:cNvPr");
      if (cNvPr) {
        const name = attr(cNvPr, "name");
        if (name) result.name = name;
      }
    }

    // Media data from a:audioFile (r:link) or p14:media (r:embed)
    const audioFileEl = findDeep(el, "a:audioFile")[0];
    const rLink = audioFileEl ? attr(audioFileEl, "r:link") : undefined;
    const p14media = !audioFileEl ? findDeep(el, "p14:media")[0] : undefined;
    const rEmbed = p14media ? attr(p14media, "r:embed") : undefined;
    const mediaRef = rLink ?? rEmbed;
    if (mediaRef) {
      const mediaPath = _ctx.resolveRelationship(mediaRef);
      if (mediaPath) {
        const data = _ctx.getRaw(mediaPath);
        if (data) result.data = data;
        result.type = mediaTypeFromPath(mediaPath, "audio");
      }
    }

    return result as AudioDescriptorOptions;
  },
};

// ── Helpers ──

function mediaTypeFromPath(path: string, kind: "video"): VideoType;
function mediaTypeFromPath(path: string, kind: "audio"): AudioType;
function mediaTypeFromPath(path: string, kind: "video" | "audio"): VideoType | AudioType {
  const ext = path.split(".").pop()?.toLowerCase() ?? "";
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
