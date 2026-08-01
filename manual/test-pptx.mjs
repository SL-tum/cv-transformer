import { readFile, writeFile } from "node:fs/promises";
import { exportPptx, importPptx } from "../dist/index.js";
const [inputPath = "manual/input.pptx", outputPath = "manual/output.pptx"] = process.argv.slice(2);
const document = importPptx(await readFile(inputPath)); let changed;
const edit = (elements) => { for (const element of elements) { if ((element.type === "shape" || element.type === "placeholder") && element.textBody) for (const paragraph of element.textBody.paragraphs) { const run = paragraph.runs.find((item) => item.type === "textRun" && item.text); if (run?.type === "textRun") { console.log(`Original: ${run.text}`); run.text = `【RDT】${run.text}`; return true; } } if (element.type === "group" && edit(element.elements)) return true; } return false; };
for (const slide of document.root.slides) if (edit(slide.shapes)) { changed = true; break; }
if (!changed) throw new Error("No editable shape text found"); await writeFile(outputPath, exportPptx(document)); console.log(`Written: ${outputPath}`);
