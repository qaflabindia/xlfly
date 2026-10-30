// Download numpy, pandas, matplotlib and their dependencies from the Pyodide release
// into vendor/pyodide, verifying each file against the SHA-256 in pyodide-lock.json.
// Self-hosting keeps the add-in independent of third-party CDNs at run time.
import { createHash } from "node:crypto";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";

const lock = JSON.parse(readFileSync("node_modules/pyodide/pyodide-lock.json", "utf8"));
const version = JSON.parse(readFileSync("node_modules/pyodide/package.json", "utf8")).version;
const source = (process.env.XLFLY_PYODIDE_SOURCE ?? `https://cdn.jsdelivr.net/pyodide/v${version}/full/`).replace(/\/?$/, "/");
const wanted = (process.env.XLFLY_PACKAGES ?? "numpy,pandas,matplotlib").split(",").map((s) => s.trim().toLowerCase());
const outDir = "vendor/pyodide";
mkdirSync(outDir, { recursive: true });

const closure = new Set();
const add = (name) => {
  const key = name.toLowerCase();
  if (closure.has(key)) return;
  const entry = lock.packages[key];
  if (!entry) throw new Error(`package ${name} is not in the Pyodide lock file`);
  closure.add(key);
  for (const dep of entry.depends ?? []) add(dep);
};
wanted.forEach(add);

for (const name of closure) {
  const { file_name: file, sha256 } = lock.packages[name];
  const target = join(outDir, file);
  if (existsSync(target) && createHash("sha256").update(readFileSync(target)).digest("hex") === sha256) continue;
  const response = await fetch(source + file);
  if (!response.ok) throw new Error(`${file}: HTTP ${response.status}`);
  const bytes = Buffer.from(await response.arrayBuffer());
  const digest = createHash("sha256").update(bytes).digest("hex");
  if (digest !== sha256) throw new Error(`${file}: checksum mismatch`);
  writeFileSync(target, bytes);
  console.log(`fetched ${file}`);
}
console.log(`${closure.size} packages in ${outDir}`);
