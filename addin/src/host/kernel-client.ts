// Host side of the kernel: owns the worker, restarts it, and enforces the context rule.

import { RpcEndpoint, RpcError, type Message, type RequestMeta } from "../shared/rpc.ts";
import type { CallItem, CallResult } from "./batcher.ts";

export interface WorkerLike {
  postMessage(message: unknown): void;
  terminate(): void;
  onmessage: ((event: { data: any }) => void) | null;
  onerror: ((event: any) => void) | null;
}

/** Excel operations Python may request. Only reachable from commands and event handlers. */
export interface HostServices {
  range_read(params: any): Promise<unknown>;
  range_write(params: any): Promise<unknown>;
  excel_call(params: any): Promise<unknown>;
  chart_native(params: any): Promise<unknown>;
  image_insert(params: any): Promise<unknown>;
  log(params: any): Promise<unknown>;
}

export interface Catalog {
  functions: Record<string, { signature: string; doc: string }>;
  commands: Record<string, { signature: string; doc: string }>;
  events: string[];
  errors: { module: string; code: string; message: string; trace?: string }[];
}

export type KernelStatus = "stopped" | "starting" | "ready" | "failed";

export interface KernelClientOptions {
  createWorker: () => WorkerLike;
  startParams: Record<string, unknown>;
  services: HostServices;
  startTimeoutMs?: number;
  functionTimeoutMs?: number;
  commandTimeoutMs?: number;
  onStatus?: (status: KernelStatus, detail?: string) => void;
  onOutput?: (stream: "stdout" | "stderr", text: string) => void;
}

const EMPTY_CATALOG: Catalog = { functions: {}, commands: {}, events: [], errors: [] };

export class KernelClient {
  status: KernelStatus = "stopped";
  catalog: Catalog = EMPTY_CATALOG;
  info: Record<string, string> | null = null;

  private readonly options: KernelClientOptions;
  private worker: WorkerLike | null = null;
  private rpc: RpcEndpoint | null = null;
  private starting: Promise<void> | null = null;
  private modules: { name: string; source: string }[] | null = null;
  private readonly commandTokens = new Set<string>();
  private commandChain: Promise<unknown> = Promise.resolve();

  constructor(options: KernelClientOptions) {
    this.options = options;
  }

  ensureStarted(): Promise<void> {
    if (this.status === "ready") return Promise.resolve();
    if (!this.starting) {
      this.starting = this.start().finally(() => {
        this.starting = null;
      });
    }
    return this.starting;
  }

  private async start(): Promise<void> {
    this.setStatus("starting");
    const worker = this.options.createWorker();
    const rpc = new RpcEndpoint((message: Message) => worker.postMessage(message));
    worker.onmessage = (event) => void rpc.handle(event.data);
    worker.onerror = (event) => this.fail(`kernel crashed: ${event?.message ?? "unknown error"}`);
    this.worker = worker;
    this.rpc = rpc;

    for (const [method, impl] of Object.entries(this.options.services)) {
      rpc.on(method, (params: unknown, meta: RequestMeta) => {
        this.guard(method, meta);
        return (impl as (p: unknown) => Promise<unknown>).call(this.options.services, params);
      });
    }
    rpc.onNotification("stdout", (p) => this.options.onOutput?.("stdout", p.text));
    rpc.onNotification("stderr", (p) => this.options.onOutput?.("stderr", p.text));

    try {
      this.info = await rpc.request("start", this.options.startParams, {
        timeoutMs: this.options.startTimeoutMs ?? 180_000,
      });
      this.setStatus("ready");
      if (this.modules) await this.sendModules(this.modules);
    } catch (error) {
      this.fail(`kernel failed to start: ${error instanceof Error ? error.message : String(error)}`);
      throw error;
    }
  }

  /** The context rule: Excel access only with a live command token. */
  private guard(method: string, meta: RequestMeta): void {
    if (!meta.context || !this.commandTokens.has(meta.context)) {
      throw new RpcError("CONTEXT", `${method} is only allowed from commands and event handlers`);
    }
  }

  async loadModules(modules: { name: string; source: string }[]): Promise<Catalog> {
    this.modules = modules;
    if (this.status !== "ready") {
      // start() sends the current modules once the kernel is up
      await this.ensureStarted();
      return this.catalog;
    }
    return this.sendModules(modules);
  }

  private async sendModules(modules: { name: string; source: string }[]): Promise<Catalog> {
    this.catalog = await this.rpc!.request<Catalog>("load_modules", { modules }, { timeoutMs: this.options.commandTimeoutMs ?? 120_000 });
    return this.catalog;
  }

  async callFunctions(batch: (CallItem & { id: string })[]): Promise<CallResult[]> {
    await this.ensureStarted();
    try {
      return await this.rpc!.request<CallResult[]>("call_functions", { batch }, {
        timeoutMs: this.options.functionTimeoutMs ?? 30_000,
      });
    } catch (error) {
      // Synchronous Python cannot be interrupted; a timeout means the kernel is stuck.
      if (error instanceof RpcError && error.code === "TIMEOUT") this.restart("worksheet function timed out");
      throw error;
    }
  }

  runCommand(name: string): Promise<{ ok: boolean; value?: unknown; error?: { code: string; message: string; trace?: string } }> {
    return this.withToken((token) =>
      this.rpc!.request("run_command", { name }, { context: token, timeoutMs: this.options.commandTimeoutMs ?? 300_000 }),
    );
  }

  dispatchEvent(event: Record<string, unknown>): Promise<{ ok: boolean; errors: unknown[] }> {
    return this.withToken((token) =>
      this.rpc!.request("event", { event }, { context: token, timeoutMs: this.options.commandTimeoutMs ?? 300_000 }),
    );
  }

  /** Commands and events run one at a time so that workbook writes never interleave. */
  private withToken<T>(run: (token: string) => Promise<T>): Promise<T> {
    const next = this.commandChain.then(async () => {
      await this.ensureStarted();
      const token = crypto.randomUUID();
      this.commandTokens.add(token);
      try {
        return await run(token);
      } catch (error) {
        if (error instanceof RpcError && error.code === "TIMEOUT") this.restart("command timed out");
        throw error;
      } finally {
        this.commandTokens.delete(token);
      }
    });
    this.commandChain = next.catch(() => undefined);
    return next;
  }

  restart(reason: string): void {
    this.stop(reason);
    this.setStatus("stopped", reason);
  }

  stop(reason = "stopped"): void {
    this.rpc?.close(reason);
    this.worker?.terminate();
    this.rpc = null;
    this.worker = null;
    this.commandTokens.clear();
    this.status = "stopped";
  }

  private fail(detail: string): void {
    this.stop(detail);
    this.setStatus("failed", detail);
  }

  private setStatus(status: KernelStatus, detail?: string): void {
    this.status = status;
    this.options.onStatus?.(status, detail);
  }
}
