import { readFileSync } from "node:fs";
import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";
import { fileURLToPath } from "node:url";
import { inflateRawSync } from "node:zlib";

import { generateOffice, parseOffice } from "../../packages/office-open/dist/index.mjs";

const root = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../..");
const outputDir = path.join(root, ".temp/cross-convert");

const groups = [
  {
    family: "document",
    source: "doc",
    fixture: "doc.b64z",
    targets: ["docx", "docm", "dotx", "dotm", "rtf", "odt"],
  },
  {
    family: "workbook",
    source: "xls",
    fixture: "xls.b64z",
    targets: ["xlsx", "xlsm", "xltx", "xltm", "ods"],
  },
  {
    family: "presentation",
    source: "ppt",
    fixture: "ppt.b64z",
    targets: ["pptx", "pptm", "potx", "potm", "odp"],
  },
];

function fixture(group) {
  const compressed = readFileSync(path.join(root, "scripts/fixtures", group.fixture));
  return new Uint8Array(inflateRawSync(Buffer.from(compressed.toString("utf8").trim(), "base64")));
}

function rtfProjection(options) {
  const allowed = [
    "sections",
    "title",
    "subject",
    "creator",
    "keywords",
    "description",
    "category",
    "appProperties",
    "footnotes",
  ];
  return Object.fromEntries(
    allowed.filter((name) => name in options).map((name) => [name, options[name]]),
  );
}

async function convertSource(family, source, data, targets) {
  const parsed = await parseOffice(data);
  const generated = [];
  for (const target of targets) {
    if (target === source) continue;
    const options = target === "rtf" ? rtfProjection(parsed.options) : { ...parsed.options };
    const output = await generateOffice(target, options, "nodebuffer");
    const outputPath = path.join(outputDir, `${family}-${source}-to-${target}.${target}`);
    await writeFile(outputPath, output);
    generated.push({ data: output, target });
  }
  return generated;
}

await mkdir(outputDir, { recursive: true });
let written = 0;
const failures = [];
for (const group of groups) {
  let firstPass;
  try {
    firstPass = await convertSource(group.family, group.source, fixture(group), group.targets);
  } catch (error) {
    failures.push(`${group.family}: ${error.message}`);
    continue;
  }
  let secondPassCount = 0;
  for (const source of firstPass) {
    try {
      const secondPass = await convertSource(
        group.family,
        source.target,
        source.data,
        group.targets,
      );
      secondPassCount += secondPass.length;
    } catch (error) {
      failures.push(`${group.family} ${source.target}: ${error.message}`);
    }
  }
  written += firstPass.length + secondPassCount;
  console.log(`${group.family}: ${firstPass.length + secondPassCount} files`);
}
console.log(`cross conversion output: ${outputDir} (${written} files)`);
if (failures.length > 0) {
  for (const failure of failures) console.error(`FAIL ${failure}`);
  process.exitCode = 1;
}
