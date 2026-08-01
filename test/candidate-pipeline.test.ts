import assert from "node:assert/strict";
import { mkdtemp, readFile, readdir, writeFile } from "node:fs/promises";
import os from "node:os";
import path from "node:path";
import test from "node:test";
import {
  loadCandidateGroundTruth,
  evaluateOperationLayoutFit,
  normalizeLlmTemplateOperation,
  recordLlmResponseBody,
  runOutputQualityRetryLoop,
  saveCandidateGroundTruth,
} from "../src/index.js";

test("output quality loop retries rejected generations and accepts the first passing output", async () => {
  const generated: number[] = [];
  const result = await runOutputQualityRetryLoop(
    async (attempt) => {
      generated.push(attempt);
      return attempt;
    },
    (attempt) => ({
      valid: attempt === 3,
      errors: attempt === 3 ? [] : [`attempt ${attempt} rejected`],
      warnings: [],
      checks: {
        opcValid: attempt === 3,
        rdtValid: attempt === 3,
        operationContentPresent: attempt === 3,
        whitespaceValid: true,
        blankTableRows: 0,
      },
    }),
    { maxAttempts: 3 },
  );
  assert.deepEqual(generated, [1, 2, 3]);
  assert.equal(result.value, 3);
  assert.equal(result.attempts, 3);
  assert.deepEqual(
    result.history.map((attempt) => attempt.valid),
    [false, false, true],
  );
});

test("LLM operation normalization removes blank lines before template execution", () => {
  assert.deepEqual(
    normalizeLlmTemplateOperation({
      op: "setList",
      targetId: "relevant-experience",
      items: ["Role A  \r\n\r\nResponsibility A\v\vResult A  ", "Role B"],
    }),
    {
      op: "setList",
      targetId: "relevant-experience",
      items: ["Role A\nResponsibility A\nResult A", "Role B"],
    },
  );
});

test("layout fit check detects content that exceeds a template section budget", () => {
  const section = {
    targetId: "summary",
    operationType: "setText" as const,
    layoutCapacity: {
      widthPt: 200,
      heightPt: 50,
      fontSizePt: 10,
      estimatedCharactersPerLine: 20,
      estimatedMaxLines: 2,
      recommendedCharacters: 34,
      recommendedItems: 1,
      sharedShapeTargets: 1,
    },
  };
  const fit = evaluateOperationLayoutFit(section, {
    op: "setText",
    targetId: "summary",
    value: "This content is deliberately long enough to require more than two rendered lines.",
  });
  assert.equal(fit.fits, false);
  assert.ok(fit.estimatedLines > 2);
  assert.match(fit.reasons.join(" "), /exceeds/);
});

test("LLM response recorder stores only the returned body in timestamped files", async () => {
  const folder = await mkdtemp(path.join(os.tmpdir(), "rdt-llm-response-"));
  const now = new Date("2026-08-01T18:30:45.123Z");
  const body = '{"schemaVersion":"1.0","operations":[]}';
  const first = await recordLlmResponseBody(body, folder, now);
  const second = await recordLlmResponseBody(body, folder, now);
  assert.equal(path.basename(first), "2026-08-01T18-30-45-123Z.json");
  assert.equal(path.basename(second), "2026-08-01T18-30-45-123Z-1.json");
  assert.equal(await readFile(first, "utf8"), `${body}\n`);
  assert.deepEqual(
    (await readdir(folder)).sort(),
    [path.basename(first), path.basename(second)].sort(),
  );
});

test("candidate PDF pipeline preserves text-layer items and stores JSON plus text", async () => {
  const folder = await mkdtemp(path.join(os.tmpdir(), "rdt-candidate-"));
  await writeFile(path.join(folder, "resume.pdf"), makeTextPdf("Ada Lovelace - Analytical Engine"));
  const groundTruth = await loadCandidateGroundTruth(folder, { candidateId: "ada" });
  assert.equal(groundTruth.documents.length, 1);
  assert.equal(groundTruth.integrity.totalPages, 1);
  assert.equal(groundTruth.integrity.pagesWithText, 1);
  assert.match(groundTruth.fullText, /Ada Lovelace/);
  assert.ok(groundTruth.documents[0]!.pages[0]!.items.length > 0);
  assert.equal(groundTruth.documents[0]!.sha256.length, 64);
  const paths = await saveCandidateGroundTruth(groundTruth, path.join(folder, "output"));
  assert.match(await readFile(paths.textPath, "utf8"), /Analytical Engine/);
  assert.equal(JSON.parse(await readFile(paths.jsonPath, "utf8")).candidateId, "ada");
});

test("candidate PDF pipeline rejects pages without a text layer", async () => {
  const folder = await mkdtemp(path.join(os.tmpdir(), "rdt-candidate-empty-"));
  await writeFile(path.join(folder, "scan.pdf"), makeTextPdf(""));
  await assert.rejects(() => loadCandidateGroundTruth(folder), /no extractable text/);
});

function makeTextPdf(text: string): Uint8Array {
  const escaped = text.replace(/([\\()])/gu, "\\$1");
  const stream = text ? `BT /F1 12 Tf 72 720 Td (${escaped}) Tj ET` : "";
  const objects = [
    "<< /Type /Catalog /Pages 2 0 R >>",
    "<< /Type /Pages /Kids [3 0 R] /Count 1 >>",
    "<< /Type /Page /Parent 2 0 R /MediaBox [0 0 612 792] /Resources << /Font << /F1 5 0 R >> >> /Contents 4 0 R >>",
    `<< /Length ${Buffer.byteLength(stream)} >>\nstream\n${stream}\nendstream`,
    "<< /Type /Font /Subtype /Type1 /BaseFont /Helvetica >>",
  ];
  let pdf = "%PDF-1.4\n";
  const offsets = [0];
  objects.forEach((object, index) => {
    offsets.push(Buffer.byteLength(pdf));
    pdf += `${index + 1} 0 obj\n${object}\nendobj\n`;
  });
  const xref = Buffer.byteLength(pdf);
  pdf += `xref\n0 ${objects.length + 1}\n0000000000 65535 f \n`;
  for (const offset of offsets.slice(1)) pdf += `${String(offset).padStart(10, "0")} 00000 n \n`;
  pdf += `trailer\n<< /Size ${objects.length + 1} /Root 1 0 R >>\nstartxref\n${xref}\n%%EOF\n`;
  return new TextEncoder().encode(pdf);
}
