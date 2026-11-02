import { copyFile, mkdir } from "node:fs/promises";
import { resolve } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(fileURLToPath(new URL("..", import.meta.url)));
const output = resolve(root, "site");
await mkdir(output, { recursive: true });
await copyFile(resolve(root, "apps/pages/index.html"), resolve(output, "index.html"));
