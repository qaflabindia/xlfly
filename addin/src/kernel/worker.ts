// Module Web Worker entry point. The host terminates this worker to stop runaway Python.

import { RpcEndpoint, type Message } from "../shared/rpc.ts";
import { startKernel } from "./kernel-core.ts";

const scope = globalThis as any;
const rpc = new RpcEndpoint((message: Message) => scope.postMessage(message));
scope.onmessage = (event: MessageEvent) => void rpc.handle(event.data);

rpc.on("start", async (config: { indexURL: string; packageBaseUrl?: string; pythonBundleUrl: string }) => {
  const { loadPyodide } = await import(/* runtime URL, not bundled */ `${config.indexURL}pyodide.mjs`);
  const response = await fetch(config.pythonBundleUrl);
  if (!response.ok) throw new Error(`cannot load Python bundle: HTTP ${response.status}`);
  const python = await response.json();
  return startKernel(rpc, loadPyodide, {
    indexURL: config.indexURL,
    packageBaseUrl: config.packageBaseUrl,
    python,
  });
});
