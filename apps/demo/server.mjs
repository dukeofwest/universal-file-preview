import { createServer } from "node:http";
import { readFile } from "node:fs/promises";
import { resolve, extname, sep } from "node:path";
import { fileURLToPath } from "node:url";

const root = resolve(fileURLToPath(new URL("../..", import.meta.url)));
const types = { ".html": "text/html; charset=utf-8", ".js": "text/javascript; charset=utf-8", ".css": "text/css; charset=utf-8", ".json": "application/json; charset=utf-8" };
const server = createServer(async (request, response) => {
  const pathname = new URL(request.url ?? "/", "http://localhost").pathname;
  let decoded;
  try { decoded = decodeURIComponent(pathname); } catch { response.writeHead(400).end("Invalid path"); return; }
  const requested = decoded === "/" ? "/apps/demo/index.html" : decoded;
  const path = resolve(root, `.${requested}`);
  if (path !== root && !path.startsWith(`${root}${sep}`)) { response.writeHead(403).end("Forbidden"); return; }
  try {
    const body = await readFile(path);
    response.writeHead(200, { "content-type": types[extname(path)] ?? "application/octet-stream", "x-content-type-options": "nosniff" });
    response.end(body);
  } catch { response.writeHead(404).end("Not found"); }
});

server.listen(4173, "127.0.0.1", () => process.stdout.write("Demo available at http://127.0.0.1:4173\n"));
