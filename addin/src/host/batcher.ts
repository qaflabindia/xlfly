// Collects worksheet-function calls made in the same tick into one kernel round trip.

export type CallItem =
  | { kind: "eval"; code: string; args: unknown[] }
  | { kind: "call"; name: string; args: unknown[] };

export type CallResult =
  | { id: string; ok: true; value: unknown }
  | { id: string; ok: false; error: { code: string; message: string } };

type Flush = (batch: (CallItem & { id: string })[]) => Promise<CallResult[]>;

export class CallBatcher {
  private queue: { item: CallItem & { id: string }; resolve: (r: CallResult) => void; reject: (e: unknown) => void }[] = [];
  private scheduled = false;
  private counter = 0;
  private readonly flushFn: Flush;
  private readonly maxBatch: number;

  constructor(flush: Flush, maxBatch = 1000) {
    this.flushFn = flush;
    this.maxBatch = maxBatch;
  }

  enqueue(item: CallItem): Promise<CallResult> {
    const id = String(++this.counter);
    return new Promise((resolve, reject) => {
      this.queue.push({ item: { ...item, id }, resolve, reject });
      if (this.queue.length >= this.maxBatch) {
        void this.flush();
      } else if (!this.scheduled) {
        this.scheduled = true;
        setTimeout(() => void this.flush(), 0);
      }
    });
  }

  private async flush(): Promise<void> {
    this.scheduled = false;
    const entries = this.queue.splice(0, this.queue.length);
    if (entries.length === 0) return;
    try {
      const results = await this.flushFn(entries.map((e) => e.item));
      const byId = new Map(results.map((r) => [r.id, r]));
      for (const entry of entries) {
        const result = byId.get(entry.item.id);
        if (result) entry.resolve(result);
        else entry.reject(new Error("kernel returned no result for call"));
      }
    } catch (error) {
      for (const entry of entries) entry.reject(error);
    }
  }
}
