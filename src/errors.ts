export type PreviewErrorCode =
  | "unsupported-format"
  | "invalid-file"
  | "resource-limit"
  | "network-error"
  | "parse-error"
  | "unsupported-feature"
  | "renderer-error"
  | "aborted";

/**
 * Error options carrying a name taken from the document (a ZIP entry or package part name) for diagnostics. Such
 * names are kept out of error messages, which the viewer shows to users: a file could otherwise put its own
 * wording into the viewer's status line. Internal helper.
 */
export function documentDetail(label: string, value: string): ErrorOptions {
  return { cause: { [label]: value.slice(0, 256) } };
}

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
