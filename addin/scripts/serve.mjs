// HTTPS static server for dist/ during development. Certificates come from office-addin-dev-certs.
import { createServer } from "node:https";
import { createReadStream, existsSync, readFileSync, statSync } from "node:fs";
import { homedir } from "node:os";
import { extname, join, normalize, resolve } from "node:path";

const root = resolve("dist");
const port = Number(process.env.PORT ?? 3000);
const certDir = join(homedir(), ".office-addin-dev-certs");
if (!existsSync(join(certDir, "localhost.crt"))) {
  console.error("No development certificate. Run: npx office-addin-dev-certs install");
  process.exit(1);
}
const types = {
  ".html": "text/html; charset=utf-8", ".js": "text/javascript", ".mjs": "text/javascript",
  ".css": "text/css", ".json": "application/json", ".wasm": "application/wasm",
  ".zip": "application/zip", ".whl": "application/zip", ".png": "image/png",
  ".xml": "application/xml", ".map": "application/json", ".tar": "application/x-tar",
};

createServer(
  { key: readFileSync(join(certDir, "localhost.key")), cert: readFileSync(join(certDir, "localhost.crt")) },
  (req, res) => {
    const path = normalize(decodeURIComponent(new URL(req.url, "https://localhost").pathname));
    const file = join(root, path.endsWith("/") ? join(path, "taskpane.html") : path);
    if (!file.startsWith(root) || !existsSync(file) || !statSync(file).isFile()) {
      res.writeHead(404).end("not found");
      return;
    }
    res.writeHead(200, {
      "Content-Type": types[extname(file)] ?? "application/octet-stream",
      "Cache-Control": "no-store",
      "Access-Control-Allow-Origin": "*",
    });
    createReadStream(file).pipe(res);
  },
).listen(port, () => console.log(`serving dist/ at https://localhost:${port}`));
