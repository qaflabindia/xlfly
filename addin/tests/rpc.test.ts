import { test } from "node:test";
import assert from "node:assert/strict";
import { RpcEndpoint, RpcError, type Message } from "../src/shared/rpc.ts";

function pair(): [RpcEndpoint, RpcEndpoint] {
  let a: RpcEndpoint, b: RpcEndpoint;
  a = new RpcEndpoint((m: Message) => queueMicrotask(() => void b.handle(structuredClone(m))));
  b = new RpcEndpoint((m: Message) => queueMicrotask(() => void a.handle(structuredClone(m))));
  return [a, b];
}

test("requests flow in both directions", async () => {
  const [host, kernel] = pair();
  kernel.on("add", ({ x, y }) => x + y);
  host.on("ping", (_p, meta) => `pong:${meta.context}`);
  assert.equal(await host.request("add", { x: 2, y: 3 }), 5);
  assert.equal(await kernel.request("ping", null, { context: "tok" }), "pong:tok");
});

test("a handler can call back across the channel while handling a request", async () => {
  const [host, kernel] = pair();
  host.on("read", () => [[1, 2]]);
  kernel.on("run", async () => (await kernel.request<number[][]>("read", {}))[0][1]);
  assert.equal(await host.request("run", {}), 2);
});

test("errors keep their code", async () => {
  const [host, kernel] = pair();
  kernel.on("boom", () => { throw new RpcError("CONTEXT", "not allowed"); });
  await assert.rejects(host.request("boom", {}), (e: any) => e.code === "CONTEXT" && e.message === "CONTEXT: not allowed");
  await assert.rejects(host.request("missing", {}), (e: any) => e.code === "METHOD_NOT_FOUND");
});

test("timeouts reject and late replies are ignored", async () => {
  const [host, kernel] = pair();
  kernel.on("slow", () => new Promise((r) => setTimeout(() => r(1), 50)));
  await assert.rejects(host.request("slow", {}, { timeoutMs: 5 }), (e: any) => e.code === "TIMEOUT");
  await new Promise((r) => setTimeout(r, 80));
  assert.equal(host.outstanding, 0);
});

test("close rejects outstanding requests", async () => {
  const [host, kernel] = pair();
  kernel.on("never", () => new Promise(() => {}));
  const pending = host.request("never", {});
  host.close("kernel died");
  await assert.rejects(pending, (e: any) => e.code === "CLOSED");
  await assert.rejects(host.request("x", {}), (e: any) => e.code === "CLOSED");
});

test("protocol version mismatch is refused", async () => {
  const sent: Message[] = [];
  const endpoint = new RpcEndpoint((m) => sent.push(m));
  await endpoint.handle({ v: 99, kind: "req", id: 1, method: "x", params: null } as any);
  assert.equal(sent.length, 1);
  assert.equal((sent[0] as any).error.code, "PROTOCOL");
});
