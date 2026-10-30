// Symmetric RPC over any message channel (Web Worker, MessagePort).
// Either side may send requests; each side numbers its own requests.

export const PROTOCOL_VERSION = 1;

export type Message =
  | { v: number; kind: "req"; id: number; method: string; params: unknown; context?: string; deadline?: number }
  | { v: number; kind: "res"; id: number; ok: true; result: unknown }
  | { v: number; kind: "res"; id: number; ok: false; error: { code: string; message: string } }
  | { v: number; kind: "note"; method: string; params: unknown };

export class RpcError extends Error {
  readonly code: string;
  constructor(code: string, message: string) {
    super(`${code}: ${message}`);
    this.code = code;
    this.name = "RpcError";
  }
}

export interface RequestMeta {
  context?: string;
  deadline?: number;
}

export type Handler = (params: any, meta: RequestMeta) => unknown | Promise<unknown>;

export interface RequestOptions {
  context?: string;
  timeoutMs?: number;
}

interface Pending {
  resolve: (value: unknown) => void;
  reject: (reason: unknown) => void;
  timer?: ReturnType<typeof setTimeout>;
}

export class RpcEndpoint {
  private nextId = 1;
  private readonly pending = new Map<number, Pending>();
  private readonly handlers = new Map<string, Handler>();
  private readonly noteHandlers = new Map<string, (params: any) => void>();
  private closed = false;
  private readonly send: (message: Message) => void;

  constructor(send: (message: Message) => void) {
    this.send = send;
  }

  on(method: string, handler: Handler): void {
    this.handlers.set(method, handler);
  }

  onNotification(method: string, handler: (params: any) => void): void {
    this.noteHandlers.set(method, handler);
  }

  request<T = unknown>(method: string, params: unknown, options: RequestOptions = {}): Promise<T> {
    if (this.closed) return Promise.reject(new RpcError("CLOSED", "channel closed"));
    const id = this.nextId++;
    const deadline = options.timeoutMs ? Date.now() + options.timeoutMs : undefined;
    return new Promise<T>((resolve, reject) => {
      const entry: Pending = { resolve: resolve as (v: unknown) => void, reject };
      if (options.timeoutMs) {
        entry.timer = setTimeout(() => {
          this.pending.delete(id);
          reject(new RpcError("TIMEOUT", `${method} exceeded ${options.timeoutMs} ms`));
        }, options.timeoutMs);
      }
      this.pending.set(id, entry);
      this.send({ v: PROTOCOL_VERSION, kind: "req", id, method, params, context: options.context, deadline });
    });
  }

  notify(method: string, params: unknown): void {
    if (!this.closed) this.send({ v: PROTOCOL_VERSION, kind: "note", method, params });
  }

  /** Feed one incoming message. */
  async handle(message: Message): Promise<void> {
    if (!message || typeof message !== "object") return;
    if (message.v !== PROTOCOL_VERSION) {
      if (message.kind === "req") {
        this.send({
          v: PROTOCOL_VERSION, kind: "res", id: message.id, ok: false,
          error: { code: "PROTOCOL", message: `unsupported protocol version ${message.v}` },
        });
      }
      return;
    }
    switch (message.kind) {
      case "res": {
        const entry = this.pending.get(message.id);
        if (!entry) return; // late reply after timeout
        this.pending.delete(message.id);
        if (entry.timer) clearTimeout(entry.timer);
        if (message.ok) entry.resolve(message.result);
        else entry.reject(new RpcError(message.error.code, message.error.message));
        return;
      }
      case "note": {
        this.noteHandlers.get(message.method)?.(message.params);
        return;
      }
      case "req": {
        const handler = this.handlers.get(message.method);
        if (!handler) {
          this.reply(message.id, new RpcError("METHOD_NOT_FOUND", message.method));
          return;
        }
        try {
          const result = await handler(message.params, { context: message.context, deadline: message.deadline });
          this.send({ v: PROTOCOL_VERSION, kind: "res", id: message.id, ok: true, result: result ?? null });
        } catch (error) {
          this.reply(message.id, error);
        }
        return;
      }
    }
  }

  /** Reject every outstanding request; used when the peer dies. */
  close(reason = "channel closed"): void {
    this.closed = true;
    for (const [id, entry] of this.pending) {
      if (entry.timer) clearTimeout(entry.timer);
      entry.reject(new RpcError("CLOSED", reason));
      this.pending.delete(id);
    }
  }

  get outstanding(): number {
    return this.pending.size;
  }

  private reply(id: number, error: unknown): void {
    const code = error instanceof RpcError ? error.code : "INTERNAL";
    const message = error instanceof RpcError
      ? error.message.slice(code.length + 2)
      : error instanceof Error ? error.message : String(error);
    this.send({ v: PROTOCOL_VERSION, kind: "res", id, ok: false, error: { code, message } });
  }
}
