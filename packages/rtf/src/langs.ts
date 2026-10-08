/**
 * RTF `\langN` language IDs (LCID) mapped to RFC 1766 tags. The table covers
 * the language IDs Word writes for locales; tags without an RTF LCID cannot
 * round-trip and are rejected on generate.
 */
export const LCID_TO_TAG: Readonly<Record<number, string>> = {
  1025: "ar-SA",
  1026: "bg-BG",
  1027: "ca-ES",
  1028: "zh-TW",
  1029: "cs-CZ",
  1030: "da-DK",
  1031: "de-DE",
  1032: "el-GR",
  1033: "en-US",
  1034: "es-ES",
  1035: "fi-FI",
  1036: "fr-FR",
  1037: "he-IL",
  1038: "hu-HU",
  1039: "is-IS",
  1040: "it-IT",
  1041: "ja-JP",
  1042: "ko-KR",
  1043: "nl-NL",
  1044: "nb-NO",
  1045: "pl-PL",
  1046: "pt-BR",
  1048: "ro-RO",
  1049: "ru-RU",
  1050: "hr-HR",
  1051: "sk-SK",
  1052: "sq-AL",
  1053: "sv-SE",
  1054: "th-TH",
  1055: "tr-TR",
  1056: "ur-PK",
  1057: "id-ID",
  1058: "uk-UA",
  1059: "be-BY",
  1060: "sl-SI",
  1061: "et-EE",
  1062: "lv-LV",
  1063: "lt-LT",
  1065: "fa-IR",
  1066: "vi-VN",
  1067: "hy-AM",
  1069: "eu-ES",
  1071: "mk-MK",
  1081: "hi-IN",
  1086: "ms-MY",
  1087: "kk-KZ",
  1089: "sw-KE",
  1091: "uz-UZ",
  1110: "gl-ES",
  2052: "zh-CN",
  2068: "nn-NO",
  2070: "pt-PT",
  2077: "sv-FI",
  3098: "sr-RS",
};

const TAG_TO_LCID = new Map(Object.entries(LCID_TO_TAG).map(([lcid, tag]) => [tag, Number(lcid)]));

export function lcidToTag(lcid: number): string | undefined {
  return LCID_TO_TAG[lcid];
}

export function tagToLcid(tag: string): number | undefined {
  return TAG_TO_LCID.get(tag);
}
