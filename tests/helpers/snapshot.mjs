// Characterization snapshots for parser models (test-only; Node built-ins only).
//
// Snapshots live in tests/snapshots/<name>.json. They must be generated on a commit whose
// src/ matches the baseline being protected:
//
//   bash:        UPDATE_SNAPSHOTS=1 npm test
//   PowerShell:  $env:UPDATE_SNAPSHOTS = "1"; npm test; Remove-Item Env:UPDATE_SNAPSHOTS
//
// Without UPDATE_SNAPSHOTS a missing snapshot fails the test, so a baseline can never be
// created silently by a later, possibly regressed, build.
import assert from "node:assert/strict";
import { createHash } from "node:crypto";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

export const snapshotDirectory = join(dirname(fileURLToPath(import.meta.url)), "..", "snapshots");
export const updatingSnapshots = process.env.UPDATE_SNAPSHOTS === "1";

/**
 * Converts a parser model into JSON-safe data without losing distinctions that plain
 * JSON.stringify would drop: explicit `undefined` properties, NaN, ±Infinity, -0, sparse
 * array holes, and binary data (recorded as byte length + SHA-256 instead of raw bytes).
 * Object keys are sorted so snapshot files are stable.
 */
export function normalize(value, path = "$") {
  if (value === undefined) return { $type: "undefined" };
  if (value === null || typeof value === "string" || typeof value === "boolean") return value;
  if (typeof value === "number") {
    if (Number.isNaN(value)) return { $type: "number", value: "NaN" };
    if (value === Infinity) return { $type: "number", value: "Infinity" };
    if (value === -Infinity) return { $type: "number", value: "-Infinity" };
    if (Object.is(value, -0)) return { $type: "number", value: "-0" };
    return value;
  }
  if (typeof value === "bigint") return { $type: "bigint", value: String(value) };
  if (ArrayBuffer.isView(value)) {
    const bytes = new Uint8Array(value.buffer, value.byteOffset, value.byteLength);
    return { $type: value.constructor.name === "Buffer" ? "Uint8Array" : value.constructor.name, byteLength: bytes.byteLength, sha256: createHash("sha256").update(bytes).digest("hex") };
  }
  if (value instanceof ArrayBuffer) return normalize(new Uint8Array(value), path);
  if (Array.isArray(value)) return Array.from({ length: value.length }, (_, index) => index in value ? normalize(value[index], `${path}[${index}]`) : { $type: "hole" });
  if (value instanceof Map) return { $type: "Map", entries: [...value].map(([key, item], index) => [normalize(key, `${path}.<key ${index}>`), normalize(item, `${path}.<value ${index}>`)]) };
  if (value instanceof Set) return { $type: "Set", values: [...value].map((item, index) => normalize(item, `${path}.<item ${index}>`)) };
  if (typeof value === "object") {
    const prototype = Object.getPrototypeOf(value), output = {};
    if (prototype !== Object.prototype && prototype !== null) output.$class = value.constructor?.name ?? "anonymous";
    for (const key of Object.keys(value).sort()) output[key] = normalize(value[key], `${path}.${key}`);
    return output;
  }
  throw new TypeError(`Cannot snapshot a ${typeof value} at ${path}.`);
}

/** Compares `value` with tests/snapshots/<name>.json, or writes it when UPDATE_SNAPSHOTS=1. */
export function matchSnapshot(name, value) {
  if (!/^[a-z0-9][a-z0-9-]*$/.test(name)) throw new Error(`Invalid snapshot name: ${name}`);
  const file = join(snapshotDirectory, `${name}.json`), actual = normalize(value);
  if (updatingSnapshots) {
    mkdirSync(snapshotDirectory, { recursive: true });
    writeFileSync(file, `${JSON.stringify(actual, null, 2)}\n`, "utf8");
    return;
  }
  if (!existsSync(file)) assert.fail(`Snapshot tests/snapshots/${name}.json is missing. Generate baselines only on an unmodified baseline commit with UPDATE_SNAPSHOTS=1 (see tests/helpers/snapshot.mjs).`);
  assert.deepStrictEqual(actual, JSON.parse(readFileSync(file, "utf8")), `Model differs from snapshot tests/snapshots/${name}.json`);
}
