import { readFile, writeFile } from "node:fs/promises";
import { loadOpcPackage, repackOfficePackage, unpackOfficePackage, validateOpcPackage } from "../dist/index.js";

const [inputPath, outputPath] = process.argv.slice(2);
if (!inputPath || !outputPath) throw new Error("Usage: node manual/test-opc.mjs input.docx output.docx");
const output = repackOfficePackage(unpackOfficePackage(await readFile(inputPath)));
const issues = validateOpcPackage(loadOpcPackage(output));
if (issues.length) throw new Error(`OPC validation failed:\n${issues.join("\n")}`);
await writeFile(outputPath, output);
console.log(`Written: ${outputPath}`);
