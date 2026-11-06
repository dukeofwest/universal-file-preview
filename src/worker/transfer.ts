/**
 * Prepares a parsed model for postMessage from the worker (design addendum A5). Internal module.
 *
 * Every Uint8Array in the model is transferred, not copied. A view that does not cover exactly its own buffer
 * (a non-zero byteOffset or a shorter byteLength, for example an image stored uncompressed inside the input
 * ZIP) or that views the input buffer is first replaced by a copy in a buffer of its own, so no larger
 * underlying buffer, and never the input, is transferred or exposed. The same view placed several times is
 * copied once, so the model keeps its shared references.
 */
function isLeaf(value: object): boolean {
  if (Array.isArray(value)) { for (const entry of value) if (typeof entry === "object" && entry !== null) return false; return true; }
  for (const key in value) { const field = (value as Record<string, unknown>)[key]; if (typeof field === "object" && field !== null) return false; }
  return true;
}

export function prepareTransfer(model: unknown, input: ArrayBuffer): ArrayBuffer[] {
  const buffers = new Set<ArrayBuffer>(), copies = new Map<Uint8Array, Uint8Array>(), seen = new Set<object>();
  const own = (bytes: Uint8Array): Uint8Array => {
    const buffer = bytes.buffer;
    if (buffer instanceof ArrayBuffer && buffer !== input && bytes.byteOffset === 0 && bytes.byteLength === buffer.byteLength) return bytes;
    let copy = copies.get(bytes);
    if (!copy) { copy = new Uint8Array(bytes.byteLength); copy.set(bytes); copies.set(bytes, copy); }
    return copy;
  };
  const stack: object[] = [];
  /** Visits one field: replaces a view that needs a copy, records its buffer, and queues objects that contain objects. */
  const visit = (container: Record<string, unknown> | unknown[], key: string | number, value: unknown) => {
    if (value instanceof Uint8Array) {
      const kept = own(value);
      if (kept !== value) (container as Record<string | number, unknown>)[key] = kept;
      buffers.add(kept.buffer as ArrayBuffer);
      return;
    }
    // Leaf objects and arrays (only primitive fields), such as cells and runs, hold no views: skipped without
    // tracking, so millions of cells do not fill the set.
    if (typeof value === "object" && value !== null && !isLeaf(value) && !seen.has(value)) { seen.add(value); stack.push(value); }
  };
  if (model instanceof Uint8Array) throw new TypeError("The model must be an object.");
  if (typeof model === "object" && model !== null) { seen.add(model); stack.push(model); }
  while (stack.length) {
    const object = stack.pop()!;
    if (Array.isArray(object)) { for (let index = 0; index < object.length; index++) visit(object, index, object[index]); continue; }
    const record = object as Record<string, unknown>;
    for (const key in record) { const value = record[key]; if (typeof value === "object" && value !== null) visit(record, key, value); }
  }
  return [...buffers];
}
