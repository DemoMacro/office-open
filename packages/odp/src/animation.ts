import { attributeString, xmlElement } from "@office-open/odf";
import type { AnimationBehavior, AnimationEntry } from "@office-open/pptx";
import type { Element } from "@office-open/xml";

import { OdpParseError } from "./error";

const ANIMATION_PATH =
  "/office:document-content/office:body/office:presentation/draw:page/animation";

const CONTAINER_NAMES = new Set(["anim:par", "anim:seq", "anim:iterate"]);

const BEHAVIOR_BY_ELEMENT = new Map<string, AnimationBehavior>([
  ["anim:set", "set"],
  ["anim:animate", "animate"],
  ["anim:animateTransform", "animateTransform"],
  ["anim:animateMotion", "animateMotion"],
  ["anim:animateColor", "animateColor"],
  ["anim:transitionFilter", "transitionFilter"],
  ["anim:audio", "audio"],
  ["anim:command", "command"],
  ["loext:animatePhysics", "animatePhysics"],
]);
const BEHAVIOR_NAMES = new Set(BEHAVIOR_BY_ELEMENT.keys());

const CONTAINER_ATTRIBUTES = [
  "presentation:group-id",
  "presentation:node-type",
  "presentation:preset-class",
  "anim:id",
  "presentation:additive",
  "presentation:preset-id",
  "presentation:preset-property",
  "presentation:preset-sub-type",
  "presentation:text-only",
  "presentation:master-element",
  "smil:accelerate",
  "smil:attributeName",
  "smil:autoReverse",
  "smil:iterate-type",
  "smil:begin",
  "smil:decelerate",
  "smil:dur",
  "smil:end",
  "smil:fill",
  "smil:id",
  "smil:repeatCount",
  "smil:repeatDur",
  "smil:restart",
  "smil:targetElement",
  "smil:sub-item",
  "anim:sub-item",
] as const;

const COMMON_BEHAVIOR_ATTRIBUTES = [
  "presentation:node-type",
  "presentation:additive",
  "presentation:preset-class",
  "presentation:preset-id",
  "presentation:preset-property",
  "presentation:preset-sub-type",
  "presentation:text-only",
  "presentation:master-element",
  "smil:accelerate",
  "smil:decelerate",
  "smil:additive",
  "smil:id",
  "smil:attributeName",
  "smil:autoReverse",
  "smil:begin",
  "smil:by",
  "smil:calcMode",
  "smil:decelerate",
  "smil:dur",
  "smil:end",
  "smil:fill",
  "smil:formula",
  "smil:from",
  "smil:id",
  "smil:repeatCount",
  "smil:repeatDur",
  "smil:restart",
  "smil:sub-item",
  "anim:sub-item",
  "smil:targetElement",
  "smil:to",
  "smil:valueType",
  "smil:values",
  "smil:keyTimes",
  "anim:color-interpolation",
  "anim:color-interpolation-direction",
] as const;

const BEHAVIOR_ATTRIBUTES: Record<string, readonly string[]> = {
  "anim:animateTransform": ["smil:type", "svg:type"],
  "anim:animateMotion": [
    "smil:origin",
    "smil:path",
    "svg:path",
    "smil:pathEditMode",
    "smil:rotate",
  ],
  "anim:animateColor": ["anim:color-interpolation", "anim:color-interpolation-direction"],
  "loext:animatePhysics": [
    "loext:density",
    "loext:bounciness",
    "loext:velocity-x",
    "loext:velocity-y",
  ],
  "anim:transitionFilter": [
    "smil:direction",
    "smil:fadeColor",
    "smil:mode",
    "smil:subtype",
    "smil:type",
  ],
  "anim:audio": ["xlink:href", "smil:audioLevel"],
  "anim:command": ["presentation:command", "smil:command"],
} satisfies Partial<Record<string, readonly string[]>>;

function animationError(element: Element, reason: string): OdpParseError {
  return new OdpParseError(reason, "content.xml", ANIMATION_PATH, element.name, reason);
}

function requireString(element: Element, name: string): string {
  const value = attributeString(element, name);
  if (value === undefined)
    throw animationError(element, `missing required animation attribute: ${name}`);
  return value;
}

function parseDuration(element: Element, value: string): number | "indefinite" {
  if (value === "indefinite") return "indefinite";
  const match = /^(\d+(?:\.\d+)?)(ms|s|min|h)$/.exec(value);
  if (!match) throw animationError(element, `invalid animation timing value: ${value}`);
  const amount = Number(match[1]);
  const unit = match[2];
  if (unit === "ms") return amount;
  if (unit === "s") return amount * 1000;
  if (unit === "min") return amount * 60_000;
  return amount * 3_600_000;
}

function durationXml(value: number | "indefinite" | string | undefined): string | undefined {
  if (value === undefined) return undefined;
  if (value === "indefinite") return value;
  if (typeof value === "string") return value;
  return `${value}ms`;
}

function assignTiming(target: AnimationEntry, element: Element, allowed: readonly string[]): void {
  for (const name of allowed) {
    const value = attributeString(element, name);
    if (value === undefined) continue;
    switch (name) {
      case "presentation:group-id":
        target.groupId = Number(value);
        break;
      case "presentation:node-type":
        target.timingNodeRole = value;
        break;
      case "presentation:preset-id":
        target.presetId = value;
        break;
      case "presentation:preset-class":
        target.presetClass = value;
        break;
      case "presentation:preset-property":
        target.presetProperty = value;
        break;
      case "presentation:text-only":
        target.presetTextOnly = value === "true";
        break;
      case "presentation:additive":
        target.presetAdditive = value as NonNullable<AnimationEntry["presetAdditive"]>;
        break;
      case "presentation:master-element":
        target.masterElement = value;
        break;
      case "anim:id":
        target.animationId = value;
        break;
      case "smil:accelerate":
        target.acceleration = Math.round(Number(value) * 100);
        break;
      case "smil:additive":
        target.additive = value as NonNullable<AnimationEntry["additive"]>;
        break;
      case "smil:decelerate":
        target.deceleration = Math.round(Number(value) * 100);
        break;
      case "smil:id":
        target.timingId = value;
        break;
      case "smil:iterate-type":
        target.iterateType = value;
        break;
      case "presentation:preset-sub-type":
        target.presetSubType = value;
        break;
      case "smil:attributeName":
        target.attributeName = value;
        break;
      case "smil:autoReverse":
        target.autoReverse = value === "true";
        break;
      case "smil:begin":
        target.begin = value === "indefinite" ? "indefinite" : value;
        break;
      case "smil:dur":
        target.duration = parseDuration(element, value);
        break;
      case "smil:end":
        target.end = value === "indefinite" ? "indefinite" : value;
        break;
      case "smil:fill":
        target.timeFill = value as NonNullable<AnimationEntry["timeFill"]>;
        break;
      case "smil:repeatCount":
        target.repeatCount = value === "indefinite" ? "indefinite" : Number(value);
        break;
      case "smil:repeatDur":
        target.repeatDuration = parseDuration(element, value);
        break;
      case "smil:restart":
        target.restart = value as NonNullable<AnimationEntry["restart"]>;
        break;
      case "smil:sub-item":
      case "anim:sub-item":
        target.targetSubItem = value;
        break;
      case "smil:targetElement":
        {
          const shapeId = /^id(\d+)$/u.exec(value);
          if (shapeId) target.shapeId = Number(shapeId[1]);
          else target.shapeName = value;
        }
        break;
      case "smil:by":
        target.animBy = value;
        break;
      case "smil:calcMode":
        target.calcMode = value as NonNullable<AnimationEntry["calcMode"]>;
        break;
      case "smil:formula":
        target.formula = value;
        break;
      case "smil:from":
        target.from = value;
        break;
      case "smil:to":
        target.to = value;
        break;
      case "smil:valueType":
        target.valueType = value as NonNullable<AnimationEntry["valueType"]>;
        break;
      case "smil:values":
        target.values = value.split(/[;,]\s*/u);
        break;
      case "smil:keyTimes":
        target.keyTimes = value.split(/[;,]\s*/u).map(Number);
        break;
      case "anim:color-interpolation":
        target.colorInterpolation = parseLiteral(element, name, value, ["rgb", "hsl"] as const);
        break;
      case "anim:color-interpolation-direction":
        target.colorInterpolationDirection = parseLiteral(element, name, value, [
          "clockwise",
          "counter-clockwise",
        ] as const);
        break;
      default:
        break;
    }
  }
}

function rejectUnknownAttributes(element: Element, allowed: readonly string[]): void {
  const unknown = Object.keys(element.attributes ?? {}).filter((name) => !allowed.includes(name));
  if (unknown.length > 0) {
    throw animationError(element, `unknown animation attributes: ${unknown.join(", ")}`);
  }
}

function parseLiteral<T extends string>(
  element: Element,
  name: string,
  value: string,
  values: readonly T[],
): T {
  const matched = values.find((item) => item === value);
  if (matched !== undefined) return matched;
  throw animationError(element, `invalid ${name}: ${value}`);
}

function requireFirst(element: Element, names: readonly string[]): string {
  for (const name of names) {
    const value = attributeString(element, name);
    if (value !== undefined) return value;
  }
  throw animationError(element, `missing required animation attribute: ${names.join(" or ")}`);
}

function rejectAnimationChildren(element: Element): void {
  if (animationElements(element).length > 0) {
    throw animationError(element, "animation behavior elements cannot contain children");
  }
}

function parseBehavior(element: Element): AnimationEntry {
  if (!element.name) throw animationError(element, "missing animation element name");
  const behavior = BEHAVIOR_BY_ELEMENT.get(element.name);
  if (behavior === undefined)
    throw animationError(element, `unsupported animation behavior: ${element.name}`);
  const specific = BEHAVIOR_ATTRIBUTES[element.name] ?? [];
  const allowed = [...COMMON_BEHAVIOR_ATTRIBUTES, ...specific];
  rejectUnknownAttributes(element, allowed);
  rejectAnimationChildren(element);

  const options: AnimationEntry = { behavior };
  assignTiming(options, element, allowed);
  const presetClass = attributeString(element, "presentation:preset-class");
  if (presetClass === "entrance" || presetClass === "exit" || presetClass === "emphasis") {
    options.class = presetClass;
  }
  const begin = attributeString(element, "smil:begin");
  if (options.timingNodeRole === "with-previous") options.trigger = "withPrevious";
  else if (options.timingNodeRole === "after-previous") options.trigger = "afterPrevious";
  else if (begin === "next") options.trigger = "onClick";

  if (behavior === "set") {
    options.setBehavior = {
      attributeName: options.attributeName ?? requireString(element, "smil:attributeName"),
      toValue: options.to ?? requireString(element, "smil:to"),
    };
  }
  if (behavior === "animateTransform") {
    const transformType =
      attributeString(element, "smil:type") ?? attributeString(element, "svg:type");
    if (transformType) {
      options.transformType = parseLiteral(element, "animation transform type", transformType, [
        "pt",
        "img",
        "scale",
        "rotate",
        "translate",
        "skewX",
        "skewY",
      ] as const);
    }
  }
  if (behavior === "animateMotion") {
    options.path = requireFirst(element, ["smil:path", "svg:path"]);
    options.pathType = "customPath";
    const origin = attributeString(element, "smil:origin");
    if (origin) options.motionOrigin = origin;
    const rotation = attributeString(element, "smil:rotate");
    if (rotation) options.motionRotation = rotation === "auto" ? "auto" : Number(rotation);
    const pathEditMode = attributeString(element, "smil:pathEditMode");
    if (pathEditMode)
      options.pathEditMode = pathEditMode as NonNullable<AnimationEntry["pathEditMode"]>;
  }
  if (behavior === "transitionFilter") {
    const type = requireString(element, "smil:type");
    options.transitionType = type;
    options.type = type as NonNullable<AnimationEntry["type"]>;
    options.transitionSubType = attributeString(element, "smil:subtype");
    options.transitionMode = attributeString(element, "smil:mode") as "in" | "out" | undefined;
    options.fadeColor = attributeString(element, "smil:fadeColor");
    options.direction = attributeString(element, "smil:direction") as
      | NonNullable<AnimationEntry["direction"]>
      | undefined;
  }
  if (behavior === "audio") {
    options.mediaType = "playAudio";
    options.audioSource = attributeString(element, "xlink:href");
    const level = attributeString(element, "smil:audioLevel");
    if (level) options.audioLevel = Number(level) * 100;
  }
  if (behavior === "command") {
    options.commandType = "call";
    options.command = requireString(element, "smil:command");
  }
  if (behavior === "animatePhysics") {
    const density = attributeString(element, "loext:density");
    const bounciness = attributeString(element, "loext:bounciness");
    const velocityX = attributeString(element, "loext:velocity-x");
    const velocityY = attributeString(element, "loext:velocity-y");
    if (
      density !== undefined ||
      bounciness !== undefined ||
      velocityX !== undefined ||
      velocityY !== undefined
    ) {
      options.physicsAnimation = {
        density: density === undefined ? undefined : Number(density),
        bounciness: bounciness === undefined ? undefined : Number(bounciness),
        velocityX: velocityX === undefined ? undefined : Number(velocityX),
        velocityY: velocityY === undefined ? undefined : Number(velocityY),
      };
    }
  }
  return options;
}

function animationElements(element: Element): Element[] {
  return (element.elements ?? []).filter((child): child is Element => child.type === "element");
}

export function parseAnimation(element: Element): AnimationEntry {
  if (!element.name) throw animationError(element, "missing animation element name");
  if (BEHAVIOR_NAMES.has(element.name)) return parseBehavior(element);
  if (!CONTAINER_NAMES.has(element.name)) {
    throw animationError(element, `unsupported animation element: ${element.name}`);
  }

  const allowed = [
    ...CONTAINER_ATTRIBUTES,
    ...(element.name === "anim:iterate" ? ["smil:iterate-type"] : []),
  ];
  rejectUnknownAttributes(element, allowed);
  const options: AnimationEntry = {
    children: animationElements(element).map(parseAnimation),
    timeNode: parseLiteral(element, "animation time node", element.name.slice("anim:".length), [
      "par",
      "seq",
      "iterate",
    ] as const),
  };
  assignTiming(options, element, allowed);
  const presetClass = attributeString(element, "presentation:preset-class");
  if (presetClass === "entrance" || presetClass === "exit" || presetClass === "emphasis") {
    options.class = presetClass;
  }
  return options;
}

export function isAnimationElement(name: string): boolean {
  return CONTAINER_NAMES.has(name) || BEHAVIOR_NAMES.has(name);
}

function animationAttributes(options: AnimationEntry): Record<string, string | number | undefined> {
  const attributes: Record<string, string | number | undefined> = {
    "presentation:preset-class": options.presetClass,
    "presentation:preset-property": options.presetProperty,
    "presentation:text-only":
      options.presetTextOnly === undefined ? undefined : String(options.presetTextOnly),
    "presentation:additive": options.presetAdditive,
    "presentation:master-element": options.masterElement,
    "anim:id": options.animationId,
    "presentation:group-id": options.groupId,
    "presentation:node-type": options.timingNodeRole,
    "presentation:preset-id": options.presetId,
    "presentation:preset-sub-type": options.presetSubType,
    "smil:id": options.timingId,
    "smil:accelerate": options.acceleration === undefined ? undefined : options.acceleration / 100,
    "smil:additive": options.additive,
    "smil:decelerate": options.deceleration === undefined ? undefined : options.deceleration / 100,
    "smil:autoReverse": options.autoReverse === undefined ? undefined : String(options.autoReverse),
    "smil:begin": durationXml(options.begin),
    "smil:dur": durationXml(options.duration),
    "smil:end": durationXml(options.end),
    "smil:fill": options.timeFill,
    "smil:repeatCount": options.repeatCount,
    "smil:repeatDur": durationXml(options.repeatDuration),
    "smil:restart": options.restart,
    "smil:sub-item": options.targetSubItem,
    "smil:targetElement":
      options.shapeId === undefined ? options.shapeName : `id${options.shapeId}`,
    "smil:attributeName": options.attributeName,
    "smil:by": options.animBy,
    "smil:calcMode": options.calcMode,
    "smil:formula": options.formula,
    "smil:from": options.from,
    "smil:to": options.to,
    "smil:valueType": options.valueType,
    "smil:values": options.values?.join(";"),
    "smil:keyTimes": options.keyTimes?.join(";"),
  };
  switch (options.behavior) {
    case "animateTransform":
      attributes["smil:type"] = options.transformType;
      break;
    case "animateMotion":
      attributes["smil:origin"] = options.motionOrigin;
      attributes["smil:path"] = options.path;
      attributes["smil:pathEditMode"] = options.pathEditMode;
      attributes["smil:rotate"] = options.motionRotation;
      break;
    case "transitionFilter":
      attributes["smil:direction"] = options.transitionDirection;
      attributes["smil:fadeColor"] = options.fadeColor;
      attributes["smil:mode"] = options.transitionMode;
      attributes["smil:subtype"] = options.transitionSubType;
      attributes["smil:type"] = options.transitionType;
      break;
    case "audio":
      attributes["xlink:href"] = options.audioSource;
      attributes["smil:audioLevel"] =
        options.audioLevel === undefined ? undefined : options.audioLevel / 100;
      break;
    case "command":
      attributes["smil:command"] = options.command;
      break;
    case "animateColor":
      attributes["anim:color-interpolation"] = options.colorInterpolation;
      attributes["anim:color-interpolation-direction"] = options.colorInterpolationDirection;
      break;
    case "animatePhysics":
      attributes["loext:density"] = options.physicsAnimation?.density;
      attributes["loext:bounciness"] = options.physicsAnimation?.bounciness;
      attributes["loext:velocity-x"] = options.physicsAnimation?.velocityX;
      attributes["loext:velocity-y"] = options.physicsAnimation?.velocityY;
      break;
    default:
      break;
  }
  return attributes;
}

export function animationXml(options: AnimationEntry): string {
  if (options.timeNode) {
    const attributes = animationAttributes(options);
    if (options.timeNode === "iterate") attributes["smil:iterate-type"] = options.iterateType;
    return xmlElement(
      `anim:${options.timeNode}`,
      animationAttributes(options),
      (options.children ?? []).map(animationXml),
    );
  }
  const name = options.behavior ? `anim:${options.behavior}` : "anim:set";
  const attributes = animationAttributes(options);
  if (options.setBehavior) {
    attributes["smil:attributeName"] = options.setBehavior.attributeName;
    attributes["smil:to"] = options.setBehavior.toValue;
  }
  return xmlElement(name, attributes);
}

export function animationsXml(options: AnimationEntry[] | undefined): string | undefined {
  if (!options?.length) return undefined;
  return options.map((option) => animationXml(option)).join("");
}
