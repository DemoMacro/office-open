# @office-open/rtf

Parse-only support for Rich Text Format (RTF). Parsed documents are projected to `DocumentOptions`.

`parseRtf` covers RTF groups and destinations, text controls, Unicode and hexadecimal escapes,
common character/paragraph formatting, font and color tables, and basic flat or nested tables.
Unsupported destinations, images, fields, headers/footers, and list metadata are skipped safely.
