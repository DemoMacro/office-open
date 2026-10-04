# @office-open/rtf

Parse-only support for Rich Text Format (RTF). Parsed documents are projected to `DocumentOptions`.

`parseRtf` covers RTF groups and destinations, text controls, Unicode and hexadecimal escapes,
common character/paragraph formatting, font and color tables, flat and nested tables, metadata,
styles, numbering, headers/footers, fields, direct hyperlinks, and shape instruction capture.
Unrecognized destinations are skipped without aborting the document.
