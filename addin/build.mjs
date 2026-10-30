// Build the add-in into dist/: bundles, static files, Python bundle, Pyodide runtime, manifest.
import * as esbuild from "esbuild";
import { cpSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { join, relative } from "node:path";

const pkg = JSON.parse(readFileSync("package.json", "utf8"));
const baseUrl = (process.env.XLFLY_BASE_URL ?? "https://localhost:3000").replace(/\/+$/, "");
const packageBaseUrl = process.env.XLFLY_PACKAGE_BASE_URL ?? "";
const out = "dist";

rmSync(out, { recursive: true, force: true });
mkdirSync(join(out, "pyodide"), { recursive: true });

const common = {
  bundle: true,
  target: ["es2020", "safari15"],
  sourcemap: true,
  logLevel: "warning",
  define: {
    XLFLY_VERSION: JSON.stringify(pkg.version),
    XLFLY_PACKAGE_BASE_URL: JSON.stringify(packageBaseUrl),
  },
};
await esbuild.build({ ...common, entryPoints: ["src/app.ts"], outfile: join(out, "app.js"), format: "iife" });
await esbuild.build({ ...common, entryPoints: ["src/kernel/worker.ts"], outfile: join(out, "kernel.js"), format: "esm" });

cpSync("src/taskpane/taskpane.html", join(out, "taskpane.html"));
cpSync("src/taskpane/taskpane.css", join(out, "taskpane.css"));
cpSync("src/functions/functions.json", join(out, "functions.json"));
cpSync("assets", join(out, "assets"), { recursive: true });

// Python package -> one JSON file the worker writes into the Pyodide file system.
const python = {};
const walk = (dir) => {
  for (const name of readdirSync(dir)) {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) { if (name !== "__pycache__") walk(path); }
    else if (name.endsWith(".py")) python[relative("python", path)] = readFileSync(path, "utf8");
  }
};
walk("python/xlfly");
writeFileSync(join(out, "python-bundle.json"), JSON.stringify(python));

// Pyodide core runtime from node_modules, plus any packages fetched into vendor/pyodide.
for (const file of ["pyodide.mjs", "pyodide.asm.mjs", "pyodide.asm.wasm", "python_stdlib.zip", "pyodide-lock.json"]) {
  cpSync(join("node_modules/pyodide", file), join(out, "pyodide", file));
}
if (existsSync("vendor/pyodide")) cpSync("vendor/pyodide", join(out, "pyodide"), { recursive: true });
else if (!packageBaseUrl) console.warn("note: vendor/pyodide is empty; run `npm run fetch:pyodide` so numpy/pandas/matplotlib can load");

writeFileSync(join(out, "manifest.xml"), readFileSync("manifest.xml", "utf8").replaceAll("{{BASE_URL}}", baseUrl));
console.log(`built ${out}/ for ${baseUrl}`);
