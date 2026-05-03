import { execSync } from "node:child_process";
/**
 * Run all demos and validate generated XML against OOXML XSD schemas.
 *
 * Usage:
 *   npx tsx scripts/validate.ts              # validate all (pptx + docx)
 *   npx tsx scripts/validate.ts pptx         # validate pptx demos only
 *   npx tsx scripts/validate.ts docx         # validate docx demos only
 *   npx tsx scripts/validate.ts slide <file> [n]  # validate specific pptx slide
 *   npx tsx scripts/validate.ts docx <file>       # validate specific docx file
 */
import * as fs from "node:fs";
import * as path from "node:path";
import { fileURLToPath } from "node:url";

import { unzipSync } from "fflate";
import { XmlDocument, XsdValidator } from "libxml2-wasm";
import { xmlRegisterFsInputProviders } from "libxml2-wasm/lib/nodejs.mjs";

const __dirname = path.dirname(fileURLToPath(import.meta.url));
const ROOT_DIR = path.resolve(__dirname, "..");
const XSD_DIR = path.resolve(__dirname, "../ooxml-schemas/legacy");

function getDemoFiles(dir: string): string[] {
    return fs
        .readdirSync(dir)
        .filter((f) => f.endsWith(".ts"))
        .sort();
}

function extractFromZip(zipPath: string, entryName: string): string {
    const buf = fs.readFileSync(zipPath);
    const files = unzipSync(buf);
    const entry = files[entryName];
    if (!entry) throw new Error(`Entry "${entryName}" not found in ${zipPath}`);
    return new TextDecoder().decode(entry);
}

function countEntries(zipPath: string, pattern: RegExp): number {
    const buf = fs.readFileSync(zipPath);
    const files = unzipSync(buf);
    return Object.keys(files).filter((k) => pattern.test(k)).length;
}

function loadValidator(xsdFile: string): XsdValidator {
    const originalCwd = process.cwd();
    process.chdir(XSD_DIR);
    const xsdBuffer = fs.readFileSync(xsdFile);
    const xsdDoc = XmlDocument.fromBuffer(xsdBuffer);
    const validator = XsdValidator.fromDoc(xsdDoc);
    xsdDoc.dispose();
    process.chdir(originalCwd);
    return validator;
}

function validateXml(validator: XsdValidator, xml: string): { pass: boolean; error?: string } {
    const doc = XmlDocument.fromString(xml);
    try {
        validator.validate(doc);
        return { pass: true };
    } catch (e) {
        if (e instanceof Error && "details" in e) {
            const details = (e as any).details as Array<{
                message: string;
                line: number;
                col: number;
            }>;
            return { pass: false, error: details.map((d) => d.message.trim()).join("\n    ") };
        }
        return { pass: false, error: (e as Error).message };
    } finally {
        doc.dispose();
    }
}

interface PackageConfig {
    name: string;
    dir: string;
    xsd: string;
    outputFile: string;
    entryPattern: RegExp;
    entryPrefix: string;
    buildCmd: string;
}

const PACKAGES: Record<string, PackageConfig> = {
    pptx: {
        name: "PPTX",
        dir: path.resolve(ROOT_DIR, "packages/pptx"),
        xsd: "pml.xsd",
        outputFile: "My Presentation.pptx",
        entryPattern: /^ppt\/slides\/slide\d+\.xml$/,
        entryPrefix: "slide",
        buildCmd: "pnpm -F @office-open/pptx build",
    },
    docx: {
        name: "DOCX",
        dir: path.resolve(ROOT_DIR, "packages/docx"),
        xsd: "wml.xsd",
        outputFile: "My Document.docx",
        entryPattern: /^word\/document\.xml$/,
        entryPrefix: "document",
        buildCmd: "pnpm -F @office-open/docx build",
    },
};

function validatePackage(pkg: PackageConfig) {
    console.log(`\n--- ${pkg.name} (${pkg.xsd}) ---`);

    const validator = loadValidator(pkg.xsd);
    const demos = getDemoFiles(path.join(pkg.dir, "demo"));
    let totalPass = 0;
    let totalFail = 0;
    const failures: string[] = [];

    for (const demo of demos) {
        try {
            execSync(`npx tsx "demo/${demo}"`, { cwd: pkg.dir, stdio: "pipe" });
        } catch {
            console.error(`  RUN FAIL: ${demo}`);
            continue;
        }

        const outputPath = path.join(pkg.dir, pkg.outputFile);
        if (!fs.existsSync(outputPath)) continue;
        const tmpFile = path.join(pkg.dir, `__validate_tmp__${path.extname(pkg.outputFile)}`);
        fs.renameSync(outputPath, tmpFile);

        let entryCount = 0;
        try {
            entryCount = countEntries(tmpFile, pkg.entryPattern);
        } catch {
            console.error(`  ZIP FAIL: ${demo}`);
            fs.unlinkSync(tmpFile);
            continue;
        }
        let demoPass = 0;
        let demoFail = 0;

        if (pkg.name === "PPTX") {
            for (let i = 1; i <= entryCount; i++) {
                const xml = extractFromZip(tmpFile, `ppt/slides/slide${i}.xml`);
                const result = validateXml(validator, xml);
                if (result.pass) {
                    demoPass++;
                    totalPass++;
                } else {
                    demoFail++;
                    totalFail++;
                    failures.push(`${demo} slide${i}: ${result.error}`);
                }
            }
        } else {
            const xml = extractFromZip(tmpFile, "word/document.xml");
            const result = validateXml(validator, xml);
            if (result.pass) {
                demoPass++;
                totalPass++;
            } else {
                demoFail++;
                totalFail++;
                failures.push(`${demo}: ${result.error}`);
            }
        }

        const label = pkg.name === "PPTX" ? ` (${entryCount} slides)` : "";
        const status = demoFail === 0 ? "PASS" : `FAIL (${demoFail}/${entryCount})`;
        console.log(`  ${status}  ${demo}${label}`);
        fs.unlinkSync(tmpFile);
    }

    validator.dispose();
    console.log(`  ${totalPass} passed, ${totalFail} failed`);
    return { totalPass, totalFail, failures };
}

function validateSingleFile(mode: string, filePath: string, slideNum?: number) {
    xmlRegisterFsInputProviders();
    const absPath = path.resolve(filePath);

    if (mode === "slide") {
        const validator = loadValidator("pml.xsd");
        const n = slideNum ?? 1;
        const xml = extractFromZip(absPath, `ppt/slides/slide${n}.xml`);
        const result = validateXml(validator, xml);
        validator.dispose();
        if (result.pass) console.log(`PASS: slide${n}`);
        else {
            console.error(`FAIL: slide${n}`);
            console.error(`  ${result.error}`);
            process.exitCode = 1;
        }
    } else if (mode === "docx") {
        const validator = loadValidator("wml.xsd");
        const xml = extractFromZip(absPath, "word/document.xml");
        const result = validateXml(validator, xml);
        validator.dispose();
        if (result.pass) console.log("PASS: document.xml");
        else {
            console.error("FAIL: document.xml");
            console.error(`  ${result.error}`);
            process.exitCode = 1;
        }
    }
}

function main() {
    const args = process.argv.slice(2);

    // Single file validation mode (requires file path argument)
    if ((args[0] === "slide" || args[0] === "docx") && args[1]) {
        xmlRegisterFsInputProviders();
        validateSingleFile(args[0], args[1], args[2] ? parseInt(args[2], 10) : undefined);
        return;
    }

    // Batch validation mode
    const targets = args.length > 0 ? args : ["pptx", "docx"];
    const invalid = targets.filter((t) => !(t in PACKAGES));
    if (invalid.length > 0) {
        console.error(`Unknown package(s): ${invalid.join(", ")}. Available: pptx, docx`);
        process.exit(1);
    }

    xmlRegisterFsInputProviders();

    // Ensure core is built
    execSync("pnpm -F @office-open/core build", { cwd: ROOT_DIR, stdio: "pipe" });

    const allFailures: string[] = [];
    let grandPass = 0;
    let grandFail = 0;

    for (const target of targets) {
        const pkg = PACKAGES[target];

        // Build the package
        execSync(pkg.buildCmd, { cwd: ROOT_DIR, stdio: "pipe" });

        const result = validatePackage(pkg);
        grandPass += result.totalPass;
        grandFail += result.totalFail;
        allFailures.push(...result.failures.map((f) => `[${target}] ${f}`));
    }

    console.log(`\n${"=".repeat(50)}`);
    console.log(`Total: ${grandPass} passed, ${grandFail} failed`);
    if (allFailures.length > 0) {
        console.log("\nFailures:");
        for (const f of allFailures) {
            console.log(`  ${f}`);
        }
        process.exitCode = 1;
    }
}

main();
