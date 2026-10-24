import { toUint8Array } from "@office-open/core";
import type {
  StylePropertyChildOptions,
  StyleResourceOptions,
  StyleSoundOptions,
} from "@office-open/odf";
import { isExternalOdfUrl, type OdfPackageFiles, xmlElement } from "@office-open/odf";
import type { AnimationEntry, TransitionOptions } from "@office-open/pptx";

import { ANIMATION_PATH } from "./animation";
import type { OdpImage } from "./drawing";
import { OdpGenerateError, OdpParseError } from "./error";

/** PPTX transition plus the transition fields ODF can express beyond OOXML. */
export interface OdpTransitionOptions extends TransitionOptions {
  filterDuration?: number | "indefinite";
  mode?: "in" | "out";
  fadeColor?: string;
  soundHref?: string;
  soundPlayFull?: boolean;
  soundXmlId?: string;
}

const TRANSITION_TYPES: Record<string, NonNullable<TransitionOptions["type"]>> = {
  barWipe: "wipe",
  checkerBoardWipe: "checker",
  ellipseWipe: "circle",
  fade: "fade",
  fanWipe: "wedge",
  fourBoxWipe: "plus",
  irisWipe: "diamond",
  dissolve: "dissolve",
  pushWipe: "push",
  randomBarWipe: "randomBar",
  slideWipe: "cover",
  pinWheelWipe: "wheel",
  waterfallWipe: "strips",
  zoom: "zoom",
};

const SMIL_TRANSITION_TYPES: Partial<Record<NonNullable<TransitionOptions["type"]>, string>> = {
  blinds: "blindsWipe",
  checker: "checkerBoardWipe",
  circle: "ellipseWipe",
  cover: "slideWipe",
  diamond: "irisWipe",
  dissolve: "dissolve",
  fade: "fade",
  plus: "fourBoxWipe",
  push: "pushWipe",
  random: "random",
  randomBar: "randomBarWipe",
  cut: "fade",
  wedge: "fanWipe",
  pull: "slideWipe",
  strips: "waterfallWipe",
  wheel: "pinWheelWipe",
  wipe: "barWipe",
  zoom: "zoom",
};

export function findTransitionEntry(entry: AnimationEntry): AnimationEntry | undefined {
  if (entry.behavior === "transitionFilter") return entry;
  return (entry.children ?? []).map(findTransitionEntry).find(Boolean);
}

export function hasTransitionOverlay(entry: AnimationEntry | undefined): boolean {
  return (
    entry?.duration === "indefinite" ||
    entry?.transitionMode !== undefined ||
    entry?.fadeColor !== undefined
  );
}

export function parseTransition(
  entries: readonly AnimationEntry[],
  transitionStyle?: StyleResourceOptions,
  packageMembers?: OdfPackageFiles,
  binaries?: Record<string, Uint8Array>,
): OdpTransitionOptions | undefined {
  const entry = entries.map(findTransitionEntry).find(Boolean);
  const transition = entry === undefined ? {} : (transitionFromEntry(entry) ?? {});
  const style = transitionStyle === undefined ? {} : transitionFromStyle(transitionStyle);
  if (Object.keys(transition).length === 0 && Object.keys(style).length === 0) return undefined;
  const result = { ...style, ...transition };
  if (packageMembers !== undefined && result.soundHref !== undefined) {
    validateTransitionSoundPackageMember(result.soundHref, packageMembers, binaries);
  }
  return result;
}

function transitionFromEntry(entry: AnimationEntry): OdpTransitionOptions | undefined {
  if (entry.transitionType === undefined)
    throw transitionParseError(entry, "missing transition type");
  const reverse = entry.transitionDirection === "reverse";
  const declaredType =
    entry.transitionType === "fade" && entry.transitionSubType === "default"
      ? "cut"
      : TRANSITION_TYPES[entry.transitionType];
  const type = declaredType === "cover" && reverse ? "pull" : declaredType;
  if (type === undefined) return undefined;
  const subtype = entry.transitionSubType;
  const stripsDirection =
    subtype === "horizontalRight"
      ? reverse
        ? "rightUp"
        : "leftDown"
      : reverse
        ? "leftUp"
        : "rightDown";
  const spokes =
    subtype === "oneBlade"
      ? 1
      : subtype === "twoBladeVertical"
        ? 2
        : subtype === "threeBlade"
          ? 3
          : subtype === "fourBlade"
            ? 4
            : subtype === "eightBlade"
              ? 8
              : undefined;
  return {
    type,
    ...(entry.duration === undefined
      ? {}
      : entry.duration === "indefinite"
        ? { filterDuration: "indefinite" }
        : { duration: entry.duration }),
    ...(entry.transitionMode === undefined ? {} : { mode: entry.transitionMode }),
    ...(entry.fadeColor === undefined ? {} : { fadeColor: entry.fadeColor }),
    ...(subtype?.includes("horizontal") || subtype === "across" ? { orient: "horizontal" } : {}),
    ...(subtype?.includes("vertical") || subtype === "down" ? { orient: "vertical" } : {}),
    ...(subtype === "fadeOverColor" || subtype === "rectangle" ? { thruBlk: true } : {}),
    ...(subtype === "fromRight" ? { direction: "left" } : {}),
    ...(subtype === "fromLeft" ? { direction: "right" } : {}),
    ...(subtype === "fromBottom" ? { direction: "up" } : {}),
    ...(subtype === "fromTop" ? { direction: "down" } : {}),
    ...(type === "strips" ? { direction: stripsDirection } : {}),
    ...(type === "wheel" && spokes !== undefined ? { spokes } : {}),
  };
}

function transitionFromStyle(resource: StyleResourceOptions): OdpTransitionOptions {
  const properties = resource.properties.find(
    (property) => property.name === "style:drawing-page-properties",
  );
  if (properties === undefined) return {};
  const speed = properties.attributes["presentation:transition-speed"];
  const duration = properties.attributes["presentation:duration"];
  const transitionType = properties.attributes["presentation:transition-type"];
  const sound = properties.children?.find((child) => child.name === "presentation:sound");
  const soundHref = sound?.attributes["xlink:href"];
  const normalizedSoundHref = soundHref === undefined ? undefined : String(soundHref);
  return {
    ...(isTransitionSpeed(speed) ? { speed } : {}),
    ...(transitionType === "manual"
      ? { advanceOnClick: true }
      : transitionType === "automatic" || transitionType === "semi-automatic"
        ? { advanceOnClick: false }
        : {}),
    ...(typeof duration === "string"
      ? { advanceAfterTime: parsePresentationDuration(duration) }
      : {}),
    ...(sound === undefined
      ? {}
      : {
          ...(normalizedSoundHref === undefined ? {} : { soundHref: normalizedSoundHref }),
          ...(sound.attributes["presentation:play-full"] === true ? { soundPlayFull: true } : {}),
          ...(sound.attributes["xml:id"] === undefined
            ? {}
            : { soundXmlId: String(sound.attributes["xml:id"]) }),
        }),
  };
}

function isTransitionSpeed(value: unknown): value is NonNullable<TransitionOptions["speed"]> {
  return value === "slow" || value === "medium" || value === "fast";
}

function parsePresentationDuration(value: string): number {
  const match = /^PT(?:(\d+)H)?(?:(\d+)M)?(?:(\d+(?:\.\d+)?)S)?$/.exec(value);
  if (
    match === null ||
    (match[1] === undefined && match[2] === undefined && match[3] === undefined)
  ) {
    throw new OdpParseError(
      `content.xml: invalid presentation duration ${value}`,
      "content.xml",
      "/office:automatic-styles/style:style/style:drawing-page-properties",
      "presentation:duration",
      "invalid-duration",
    );
  }
  return (
    Number(match[1] ?? 0) * 3_600_000 +
    Number(match[2] ?? 0) * 60_000 +
    Number(match[3] ?? 0) * 1000
  );
}

function transitionParseError(entry: AnimationEntry, reason: string): OdpParseError {
  return new OdpParseError(
    `content.xml: ${reason}`,
    "content.xml",
    ANIMATION_PATH,
    entry.behavior,
    reason,
  );
}

function validateTransitionSoundPackageMember(
  href: string,
  packageMembers: OdfPackageFiles,
  binaries?: Record<string, Uint8Array>,
): void {
  const path = packageRelativeSoundPath(href);
  if (path === undefined || packageMembers[path] !== undefined || binaries?.[path] !== undefined)
    return;
  throw new OdpParseError(
    `content.xml: missing transition sound package member ${path}`,
    "content.xml",
    "/office:automatic-styles/style:style/style:drawing-page-properties/presentation:sound",
    "presentation:sound",
    "missing-transition-sound",
  );
}

function packageRelativeSoundPath(href: string): string | undefined {
  if (isExternalOdfUrl(href)) return undefined;
  const normalized = href.replace(/^\/+/, "").replace(/^\.\//, "");
  try {
    return decodeURIComponent(normalized.replace(/%(?![0-9A-Fa-f]{2})/g, "%25"));
  } catch (cause) {
    throw new OdpParseError(
      `content.xml: invalid transition sound href ${href}`,
      "content.xml",
      "/office:automatic-styles/style:style/style:drawing-page-properties/presentation:sound",
      "presentation:sound",
      "invalid-transition-sound-href",
      { cause },
    );
  }
}

export function requiresTransitionStyle(options: OdpTransitionOptions): boolean {
  return (
    options.speed !== undefined ||
    options.advanceOnClick !== undefined ||
    options.advanceAfterTime !== undefined ||
    options.startSound !== undefined ||
    options.soundHref !== undefined ||
    options.soundPlayFull !== undefined ||
    options.soundXmlId !== undefined ||
    options.stopPreviousSound === true
  );
}

export function validateOdpTransition(options: OdpTransitionOptions): void {
  if (options.stopPreviousSound === true) {
    throw transitionGenerateError(
      "ODP cannot represent stopping the previous transition sound",
      "presentation:sound",
      "unsupported-transition-stop-previous-sound",
    );
  }
  if (options.startSound !== undefined && options.soundHref !== undefined) {
    throw transitionGenerateError(
      "ODP transition cannot have both embedded and linked sound",
      "presentation:sound",
      "ambiguous-transition-sound",
    );
  }
  if (
    options.startSound === undefined &&
    options.soundHref === undefined &&
    (options.soundPlayFull !== undefined || options.soundXmlId !== undefined)
  ) {
    throw transitionGenerateError(
      "ODP transition sound metadata requires embedded or linked sound",
      "presentation:sound",
      "transition-sound-without-carrier",
    );
  }
  const startSound = options.startSound;
  if (
    startSound !== undefined &&
    (startSound.name !== undefined || startSound.loop !== undefined)
  ) {
    throw transitionGenerateError(
      "ODP cannot represent transition sound metadata",
      "presentation:sound",
      "unsupported-transition-sound-metadata",
    );
  }
}

export function materializeOdpTransitionSound(
  options: OdpTransitionOptions,
  mediaFiles: OdpImage[],
  occupiedPaths: Set<string> = new Set(),
): OdpTransitionOptions {
  validateOdpTransition(options);
  if (options.soundHref !== undefined || options.startSound === undefined) return options;
  const { startSound, ...linkedTransition } = options;
  const data = toUint8Array(startSound.data);
  const existing = mediaFiles.find(
    (media) =>
      media.path.startsWith("Media/transition") &&
      media.path.endsWith(`.${startSound.type}`) &&
      sameBytes(media.data, data),
  );
  if (existing !== undefined) return { ...linkedTransition, soundHref: existing.path };
  let suffix = mediaFiles.length + 1;
  let href = `Media/transition${suffix}.${startSound.type}`;
  while (occupiedPaths.has(href)) href = `Media/transition${++suffix}.${startSound.type}`;
  occupiedPaths.add(href);
  mediaFiles.push({ path: href, data });
  return { ...linkedTransition, soundHref: href };
}

function sameBytes(left: Uint8Array, right: Uint8Array): boolean {
  return left.length === right.length && left.every((byte, index) => byte === right[index]);
}

type OdpDrawingPageProperties = Extract<
  StyleResourceOptions["properties"][number],
  { name: "style:drawing-page-properties" }
>;

type OdpSoundProperties = StyleSoundOptions;

export function applyTransitionToStyleResource(
  resource: StyleResourceOptions,
  options: OdpTransitionOptions,
): void {
  validateOdpTransition(options);
  const index = resource.properties.findIndex(
    (property) => property.name === "style:drawing-page-properties",
  );
  const existing =
    index === -1 ? undefined : (resource.properties[index] as OdpDrawingPageProperties);
  const property = drawingPageProperties(existing, options);
  resource.properties =
    index === -1
      ? [...resource.properties, property]
      : resource.properties.map((item, itemIndex) => (itemIndex === index ? property : item));
}

function drawingPageProperties(
  existing: OdpDrawingPageProperties | undefined,
  options: OdpTransitionOptions,
): OdpDrawingPageProperties {
  const attributes = { ...existing?.attributes };
  delete attributes["presentation:transition-style"];
  delete attributes["presentation:transition-type"];
  delete attributes["presentation:transition-speed"];
  delete attributes["presentation:duration"];
  attributes["presentation:transition-type"] =
    options.advanceOnClick === false
      ? "automatic"
      : options.advanceAfterTime !== undefined
        ? "semi-automatic"
        : "manual";
  if (options.speed !== undefined) attributes["presentation:transition-speed"] = options.speed;
  if (options.advanceAfterTime !== undefined)
    attributes["presentation:duration"] = presentationDurationXml(options.advanceAfterTime);
  const children: StylePropertyChildOptions[] = (existing?.children ?? []).filter(
    (child) => child.name !== "presentation:sound",
  );
  if (options.soundHref !== undefined || options.startSound !== undefined) {
    const sound: OdpSoundProperties = {
      name: "presentation:sound",
      attributes: {
        "xlink:type": "simple",
        ...(options.soundHref === undefined ? {} : { "xlink:href": options.soundHref }),
        ...(options.soundPlayFull === undefined
          ? {}
          : { "presentation:play-full": options.soundPlayFull }),
        ...(options.soundXmlId === undefined ? {} : { "xml:id": options.soundXmlId }),
      },
    };
    children.push(sound);
  }
  return { name: "style:drawing-page-properties", attributes, children };
}

function transitionGenerateError(
  message: string,
  elementName: string,
  reason: string,
): OdpGenerateError {
  return new OdpGenerateError(
    message,
    "content.xml",
    "/office:automatic-styles/style:style/style:drawing-page-properties/presentation:sound",
    elementName,
    reason,
  );
}

export function transitionDrawingPageXml(name: string, options: OdpTransitionOptions): string {
  validateOdpTransition(options);
  const attributes = {
    "presentation:transition-type":
      options.advanceOnClick === false
        ? "automatic"
        : options.advanceAfterTime !== undefined
          ? "semi-automatic"
          : "manual",
    "presentation:transition-speed": options.speed,
    "presentation:duration":
      options.advanceAfterTime === undefined
        ? undefined
        : presentationDurationXml(options.advanceAfterTime),
  };
  const children: string[] = [];
  if (options.soundHref !== undefined || options.startSound !== undefined) {
    children.push(
      xmlElement("presentation:sound", {
        "xlink:type": "simple",
        "xlink:href": options.soundHref,
        ...(options.soundPlayFull === undefined
          ? {}
          : { "presentation:play-full": options.soundPlayFull }),
        ...(options.soundXmlId === undefined ? {} : { "xml:id": options.soundXmlId }),
      }),
    );
  }
  return xmlElement("style:style", { "style:name": name, "style:family": "drawing-page" }, [
    xmlElement("style:drawing-page-properties", attributes, children),
  ]);
}

export function transitionXml(options: OdpTransitionOptions, _mediaFiles: OdpImage[]): string {
  validateOdpTransition(options);
  const smilType = options.type === undefined ? undefined : SMIL_TRANSITION_TYPES[options.type];
  if (smilType === undefined) {
    throw new OdpGenerateError(
      `ODP cannot represent slide transition type: ${options.type}`,
      "content.xml",
      "/office:document-content/office:body/office:presentation/draw:page/anim:transitionFilter",
      options.type ?? "(missing)",
      "unsupported-transition-type",
    );
  }
  const reverse =
    options.type === "pull" || options.direction === "leftUp" || options.direction === "rightUp";
  const direction =
    options.type === "strips" ? (options.direction ?? "leftDown") : options.direction;
  const subtype =
    options.type === "fade"
      ? options.thruBlk
        ? "fadeOverColor"
        : "crossfade"
      : options.type === "cut"
        ? "default"
        : options.type === "wheel"
          ? (
              {
                1: "oneBlade",
                2: "twoBladeVertical",
                3: "threeBlade",
                4: "fourBlade",
                8: "eightBlade",
              } as Record<number, string>
            )[options.spokes ?? 4]
          : options.type === "strips"
            ? direction === "leftUp" || direction === "rightDown"
              ? "horizontalLeft"
              : "horizontalRight"
            : options.orient !== undefined
              ? options.orient
              : direction === "left"
                ? "fromRight"
                : direction === "right"
                  ? "fromLeft"
                  : direction === "up"
                    ? "fromBottom"
                    : direction === "down"
                      ? "fromTop"
                      : undefined;
  const filter = xmlElement("anim:transitionFilter", {
    "smil:type": smilType,
    ...(reverse ? { "smil:direction": "reverse" } : {}),
    ...(subtype === undefined ? {} : { "smil:subtype": subtype }),
    ...(options.filterDuration === undefined && options.duration === undefined
      ? {}
      : { "smil:dur": durationXml(options.filterDuration ?? options.duration) }),
    ...(options.mode === undefined ? {} : { "smil:mode": options.mode }),
    ...(options.fadeColor === undefined ? {} : { "smil:fadeColor": options.fadeColor }),
  });
  return xmlElement("anim:par", { "presentation:node-type": "timing-root" }, [filter]);
}

function presentationDurationXml(value: number): string {
  const seconds = value / 1000;
  return `PT${Number.isInteger(seconds) ? seconds : seconds.toFixed(3)}S`;
}

function durationXml(value: number | "indefinite" | undefined): string | undefined {
  if (value === undefined) return undefined;
  return value === "indefinite" ? value : `${value}ms`;
}
