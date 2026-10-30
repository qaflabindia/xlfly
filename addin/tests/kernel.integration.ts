// Runs the real Pyodide kernel in-process through KernelClient, with Excel replaced by fakes.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readdirSync, readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, join } from "node:path";
import { loadPyodide } from "pyodide";
import { RpcEndpoint, type Message } from "../src/shared/rpc.ts";
import { startKernel } from "../src/kernel/kernel-core.ts";
import { KernelClient, type HostServices, type WorkerLike } from "../src/host/kernel-client.ts";

const indexURL = dirname(createRequire(import.meta.url).resolve("pyodide/package.json")) + "/";
const python: Record<string, string> = {};
for (const f of readdirSync("python/xlfly")) if (f.endsWith(".py")) python[`xlfly/${f}`] = readFileSync(join("python/xlfly", f), "utf8");

function inProcessWorker(): WorkerLike {
  const worker: WorkerLike = {
    onmessage: null,
    onerror: null,
    postMessage: (m) => queueMicrotask(() => void kernel.handle(structuredClone(m) as Message)),
    terminate: () => {},
  };
  const kernel = new RpcEndpoint((m) => queueMicrotask(() => worker.onmessage?.({ data: structuredClone(m) })));
  kernel.on("start", () => startKernel(kernel, loadPyodide as any, { indexURL, python }));
  return worker;
}

const sheet = new Map<string, unknown>([["A1", 10], ["A2", 20]]);
const writes: unknown[] = [];
const services: HostServices = {
  range_read: async ({ address }) => ({ values: [[sheet.get("A1")], [sheet.get("A2")]], address }),
  range_write: async ({ address, values }) => { writes.push({ address, values }); return { address: "B1:B2" }; },
  excel_call: async ({ name, args }) => ({ value: name === "SUM" ? (args as number[]).reduce((a, b) => a + b, 0) : null }),
  chart_native: async ({ kind }) => ({ name: `Chart ${kind}` }),
  image_insert: async ({ name }) => ({ name }),
  log: async () => null,
};

const SOURCE = `
import xlfly

@xlfly.func
def scale(values, k):
    return [[v * k] for row in values for v in row]

@xlfly.func
def peek():
    coroutine = xlfly.Range("A1").read()
    coroutine.send(None)  # the first step raises ContextError before any host request

@xlfly.command
async def report():
    values = await xlfly.Range("Data!A1:A2").read()
    total = await xlfly.call("SUM", *[r[0] for r in values])
    await xlfly.Range("B1").write([[v[0] * 2] for v in values])
    name = await xlfly.chart("column", "Data!A1:A2", "Data!D2", title="Values")
    print("report done", total)
    return [total, name]

@xlfly.on("sheet_change")
async def changed(event):
    await xlfly.log("changed " + event["address"])
`;

test("Pyodide kernel end to end", { timeout: 120_000 }, async () => {
  const output: string[] = [];
  const client = new KernelClient({
    createWorker: inProcessWorker,
    startParams: {},
    services,
    onOutput: (_s, text) => output.push(text),
  });

  const catalog = await client.loadModules([{ name: "main", source: SOURCE }]);
  assert.deepEqual(catalog.errors, []);
  assert.deepEqual(Object.keys(catalog.functions), ["PEEK", "SCALE"]);
  assert.match(client.info!.python, /^3\./);

  const results = await client.callFunctions([
    { id: "1", kind: "call", name: "scale", args: [[[1], [2]], 3] },
    { id: "2", kind: "eval", code: "sum(r[0] for r in arg1)", args: [[[1], [2], [3]]] },
    { id: "3", kind: "call", name: "peek", args: [] },
  ]);
  assert.deepEqual(results[0], { id: "1", ok: true, value: [[3], [6]] });
  assert.deepEqual(results[1], { id: "2", ok: true, value: 6 });
  assert.equal(results[2].ok, false);
  assert.match((results[2] as any).error.message, /ContextError/);
  assert.equal(writes.length, 0);

  const command = await client.runCommand("report");
  assert.deepEqual(command, { ok: true, value: [[30], ["Chart column"]] });
  assert.deepEqual(writes, [{ address: "B1", values: [[20], [40]] }]);
  assert.ok(output.some((line) => line.includes("report done 30")));

  const event = await client.dispatchEvent({ type: "sheet_change", address: "C3" });
  assert.deepEqual(event, { ok: true, errors: [] });

  client.stop();
});
