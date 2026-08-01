import { mkdir, writeFile } from "node:fs/promises";
import path from "node:path";

export const DEFAULT_LLM_RESPONSE_DIRECTORY = "manual/llm-responses";

export async function recordLlmResponseBody(
  body: string,
  directory = process.env.LLM_RESPONSE_DIR ?? DEFAULT_LLM_RESPONSE_DIRECTORY,
  now = new Date(),
): Promise<string> {
  await mkdir(directory, { recursive: true });
  const timestamp = now.toISOString().replaceAll(":", "-").replaceAll(".", "-");
  const basePath = path.join(directory, `${timestamp}.json`);
  const content = body.trimEnd() + "\n";

  for (let sequence = 0; ; sequence++) {
    const filePath =
      sequence === 0 ? basePath : path.join(directory, `${timestamp}-${sequence}.json`);
    try {
      await writeFile(filePath, content, { encoding: "utf8", flag: "wx" });
      return filePath;
    } catch (error) {
      if (isFileExistsError(error)) continue;
      throw error;
    }
  }
}

function isFileExistsError(error: unknown): error is NodeJS.ErrnoException {
  return error instanceof Error && "code" in error && error.code === "EEXIST";
}
