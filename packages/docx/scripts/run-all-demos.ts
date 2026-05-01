import fs from "fs";
import path from "path";

import { $ } from "execa";

const dir = path.resolve(import.meta.dirname, "../demo");
const fileNames = fs.readdirSync(dir);
const getFileNumber = (file: string): number => Number(file.split("-")[0]);

const demoFiles = fileNames
    .filter((f) => f.endsWith(".ts") && !isNaN(getFileNumber(f)))
    .toSorted((a, b) => getFileNumber(a) - getFileNumber(b));

let passed = 0;
let failed = 0;

for (const file of demoFiles) {
    const num = getFileNumber(file);
    try {
        console.log(`\n[${num}] Running ${file}...`);
        await $`tsx ${path.join(dir, file)}`;
        console.log(`[${num}] OK`);
        passed++;
    } catch (err) {
        console.error(`[${num}] FAILED: ${(err as Error).message}`);
        failed++;
    }
}

console.log(`\n${passed} passed, ${failed} failed, ${passed + failed} total`);
if (failed > 0) process.exit(1);
