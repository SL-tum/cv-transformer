import { readFile, writeFile } from "node:fs/promises";
import { exportDocx, importDocx } from "../dist/index.js";
const [inputPath = "manual/input.docx", outputPath = "manual/output.docx"] = process.argv.slice(2);
const document = importDocx(await readFile(inputPath)); let changed;
for (const section of document.root.sections) { for (const block of section.blocks) { if (block.type !== "paragraph" || !("runs" in block)) continue; const run = block.runs.find((item) => item.type === "textRun" && item.text); if (run?.type === "textRun") { console.log(`Original: ${run.text}`); run.text = `【RDT】${run.text}`; changed = true; break; } } if (changed) break; }
if (!changed) throw new Error("No editable body text found"); await writeFile(outputPath, exportDocx(document)); console.log(`Written: ${outputPath}`);
