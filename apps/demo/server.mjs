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

const host = process.env.HOST?.trim() || "127.0.0.1";
const port = Number(process.env.PORT ?? 4173);
if (!Number.isInteger(port) || port < 1 || port > 65535) throw new RangeError("PORT must be an integer from 1 to 65535.");
server.listen(port, host, () => process.stdout.write(`Demo available at http://${host}:${port}\n`));
