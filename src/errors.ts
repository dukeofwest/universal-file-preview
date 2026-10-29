export type PreviewErrorCode =
  | "unsupported-format"
  | "invalid-file"
  | "resource-limit"
  | "network-error"
  | "parse-error"
  | "unsupported-feature"
  | "renderer-error"
  | "aborted";

export class PreviewError extends Error {
  constructor(
    message: string,
    readonly code: PreviewErrorCode,
    options?: ErrorOptions,
  ) {
    super(message, options);
    this.name = "PreviewError";
  }
}
