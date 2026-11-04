import test from "node:test";
import assert from "node:assert/strict";
import { execFileSync } from "node:child_process";
import { existsSync } from "node:fs";
import { dirname, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = join(dirname(fileURLToPath(import.meta.url)), "..");
const tsc = join(root, "node_modules", "typescript", "bin", "tsc");

test("extended spreadsheet model types compile with the project's strict settings", { skip: !existsSync(tsc) && "typescript (devDependency) is not installed" }, () => {
  // Same strictness as tsconfig.json; the check file is compiled on its own, without emitting.
  const args = [tsc, "--noEmit", "--pretty", "false", "--target", "ES2022", "--module", "NodeNext", "--moduleResolution", "NodeNext", "--lib", "ES2022,DOM,DOM.Iterable",
    "--strict", "--noUncheckedIndexedAccess", "--exactOptionalPropertyTypes", "--noImplicitOverride", "--useUnknownInCatchVariables", "--skipLibCheck",
    join(root, "tests", "types", "model-extensions.check.ts")];
  let output = "";
  try { output = execFileSync(process.execPath, args, { cwd: root, encoding: "utf8" }); }
  catch (error) { assert.fail(`tsc reported errors:\n${error.stdout ?? ""}${error.stderr ?? ""}`); }
  assert.equal(output.trim(), "");
});
