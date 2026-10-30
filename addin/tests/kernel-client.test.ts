import { test } from "node:test";
import assert from "node:assert/strict";
import { RpcEndpoint, type Message } from "../src/shared/rpc.ts";
import { KernelClient, type HostServices, type WorkerLike } from "../src/host/kernel-client.ts";
import { CallBatcher } from "../src/host/batcher.ts";

/** An in-process fake worker whose kernel side is scripted by the test. */
function fakeWorkerFactory(setup: (rpc: RpcEndpoint) => void) {
  const created: { terminated: boolean }[] = [];
  const factory = (): WorkerLike => {
    const state = { terminated: false };
    created.push(state);
    const worker: WorkerLike = {
      onmessage: null,
      onerror: null,
      postMessage: (m) => queueMicrotask(() => { if (!state.terminated) void kernel.handle(structuredClone(m) as Message); }),
      terminate: () => { state.terminated = true; },
    };
    const kernel = new RpcEndpoint((m) => queueMicrotask(() => { if (!state.terminated) worker.onmessage?.({ data: structuredClone(m) }); }));
    kernel.on("start", () => ({ python: "test" }));
    kernel.on("load_modules", () => ({ functions: {}, commands: {}, events: [], errors: [] }));
    setup(kernel);
    return worker;
  };
  return { factory, created };
}

const calls: string[] = [];
const services: HostServices = {
  range_read: async () => { calls.push("read"); return { values: [[1]] }; },
  range_write: async () => ({ address: "A1" }),
  excel_call: async () => ({ value: 1 }),
  chart_native: async () => ({ name: "c" }),
  image_insert: async () => ({ name: "i" }),
  log: async () => null,
};

test("Excel access is refused without a command token and allowed with one", async () => {
  const { factory } = fakeWorkerFactory((rpc) => {
    rpc.on("call_functions", async () => {
      await rpc.request("range_read", { address: "A1" });  // no context: must be refused
      return [];
    });
    rpc.on("run_command", async (_p, meta) => rpc.request("range_read", { address: "A1" }, { context: meta.context }));
  });
  const client = new KernelClient({ createWorker: factory, startParams: {}, services });
  await assert.rejects(client.callFunctions([]), (e: any) => e.code === "CONTEXT");
  calls.length = 0;
  assert.deepEqual(await client.runCommand("x"), { values: [[1]] });
  assert.deepEqual(calls, ["read"]);
});

test("a stale token from a finished command is refused", async () => {
  let saved = "";
  const { factory } = fakeWorkerFactory((rpc) => {
    rpc.on("run_command", (_p, meta) => { saved = meta.context!; return { ok: true }; });
    rpc.on("call_functions", async () => rpc.request("range_read", {}, { context: saved }));
  });
  const client = new KernelClient({ createWorker: factory, startParams: {}, services });
  await client.runCommand("x");
  await assert.rejects(client.callFunctions([]), (e: any) => e.code === "CONTEXT");
});

test("a stuck worksheet function restarts the kernel and modules are reloaded", async () => {
  let loads = 0;
  const { factory, created } = fakeWorkerFactory((rpc) => {
    rpc.on("load_modules", () => { loads++; return { functions: {}, commands: {}, events: [], errors: [] }; });
    rpc.on("call_functions", () => new Promise(() => {}));
  });
  const client = new KernelClient({ createWorker: factory, startParams: {}, services, functionTimeoutMs: 20 });
  await client.loadModules([{ name: "m", source: "" }]);
  await assert.rejects(client.callFunctions([]), (e: any) => e.code === "TIMEOUT");
  assert.equal(created[0].terminated, true);
  await client.ensureStarted();
  assert.equal(created.length, 2);
  assert.equal(loads, 2);
});

test("commands run one at a time", async () => {
  let active = 0, peak = 0;
  const { factory } = fakeWorkerFactory((rpc) => {
    rpc.on("run_command", async () => {
      active++; peak = Math.max(peak, active);
      await new Promise((r) => setTimeout(r, 5));
      active--;
      return { ok: true };
    });
  });
  const client = new KernelClient({ createWorker: factory, startParams: {}, services });
  await Promise.all([client.runCommand("a"), client.runCommand("b"), client.runCommand("c")]);
  assert.equal(peak, 1);
});

test("calls in one tick are sent as one batch", async () => {
  const batches: number[] = [];
  const batcher = new CallBatcher(async (batch) => {
    batches.push(batch.length);
    return batch.map((b) => ({ id: b.id, ok: true as const, value: b.kind }));
  });
  const results = await Promise.all([
    batcher.enqueue({ kind: "eval", code: "1", args: [] }),
    batcher.enqueue({ kind: "call", name: "F", args: [] }),
    batcher.enqueue({ kind: "eval", code: "2", args: [] }),
  ]);
  assert.deepEqual(batches, [3]);
  assert.deepEqual(results.map((r) => (r.ok ? r.value : null)), ["eval", "call", "eval"]);
});
