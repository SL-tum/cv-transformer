import type { OutputQualityReport } from "./output-quality.js";

export interface OutputRetryAttempt {
  attempt: number;
  valid: boolean;
  errors: string[];
}

export interface OutputRetryResult<T> {
  value: T;
  quality: OutputQualityReport;
  attempts: number;
  history: OutputRetryAttempt[];
}

export async function runOutputQualityRetryLoop<T>(
  generate: (attempt: number, previousErrors: string[]) => Promise<T>,
  validate: (value: T) => OutputQualityReport | Promise<OutputQualityReport>,
  options: { maxAttempts?: number } = {},
): Promise<OutputRetryResult<T>> {
  const maxAttempts = options.maxAttempts ?? 3;
  if (!Number.isInteger(maxAttempts) || maxAttempts < 1)
    throw new RangeError("maxAttempts must be a positive integer");
  const history: OutputRetryAttempt[] = [];
  let previousErrors: string[] = [];
  for (let attempt = 1; attempt <= maxAttempts; attempt++) {
    try {
      const value = await generate(attempt, previousErrors);
      const quality = await validate(value);
      history.push({ attempt, valid: quality.valid, errors: quality.errors });
      if (quality.valid) return { value, quality, attempts: attempt, history };
      previousErrors = quality.errors;
    } catch (error) {
      previousErrors = [error instanceof Error ? error.message : String(error)];
      history.push({ attempt, valid: false, errors: previousErrors });
    }
  }
  throw new Error(
    `Generated output failed quality validation after ${maxAttempts} attempts: ${previousErrors.join("; ")}`,
  );
}
