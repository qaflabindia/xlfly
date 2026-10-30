// Shared-runtime entry point: custom functions, events and the task pane share this one page.

/* global Office, Excel, CustomFunctions */

import { trustKey, type CodeDocument, type Module } from "./shared/codestore-format.ts";
import { wireToMatrix, argToWire, type Wire } from "./shared/values.ts";
import { CallBatcher, type CallItem } from "./host/batcher.ts";
import { KernelClient, type KernelStatus, type WorkerLike } from "./host/kernel-client.ts";
import { createExcelServices, WriteTracker } from "./host/excel-services.ts";
import { emptyDocument, isTrusted, loadCode, saveCode, setTrusted } from "./host/office-store.ts";
import { bindTaskpane } from "./taskpane/taskpane.ts";

declare const XLFLY_VERSION: string;
declare const XLFLY_PACKAGE_BASE_URL: string;

type Listener = () => void;

export class App {
  readonly version = XLFLY_VERSION;
  doc: CodeDocument | null = null;
  trusted = false;
  kernelStatus: KernelStatus = "stopped";
  kernelDetail = "";
  readonly logLines: string[] = [];

  readonly kernel: KernelClient;
  private readonly batcher: CallBatcher;
  private readonly tracker = new WriteTracker();
  private readonly listeners = new Set<Listener>();
  private workbookOpenSent = false;

  constructor() {
    const base = new URL(".", location.href);
    this.kernel = new KernelClient({
      createWorker: () => new Worker(new URL("kernel.js", base), { type: "module" }) as unknown as WorkerLike,
      startParams: {
        indexURL: new URL("pyodide/", base).href,
        packageBaseUrl: XLFLY_PACKAGE_BASE_URL || undefined,
        pythonBundleUrl: new URL("python-bundle.json", base).href,
      },
      services: createExcelServices(this.tracker, (line) => this.log(line)),
      onStatus: (status, detail) => {
        this.kernelStatus = status;
        this.kernelDetail = detail ?? "";
        if (detail) this.log(`kernel ${status}: ${detail}`);
        this.changed();
      },
      onOutput: (_stream, text) => this.log(text),
    });
    this.batcher = new CallBatcher((batch) => this.kernel.callFunctions(batch));
  }

  async init(): Promise<void> {
    try {
      this.doc = await loadCode();
    } catch (error) {
      this.log(`cannot read workbook code: ${message(error)}`);
    }
    await this.refreshTrust();
    await this.registerEvents();
  }

  subscribe(listener: Listener): () => void {
    this.listeners.add(listener);
    return () => this.listeners.delete(listener);
  }

  log(line: string): void {
    this.logLines.push(`${new Date().toLocaleTimeString()}  ${line}`);
    if (this.logLines.length > 500) this.logLines.splice(0, this.logLines.length - 500);
    this.changed();
  }

  // ------------------------------------------------------------------ trust

  private async currentKey(): Promise<string | null> {
    return this.doc ? trustKey(this.doc, Office.context.document.url ?? "") : null;
  }

  private async refreshTrust(): Promise<void> {
    const key = await this.currentKey();
    this.trusted = key !== null && (await isTrusted(key));
    this.changed();
    if (this.trusted) await this.reload();
  }

  async trustWorkbook(): Promise<void> {
    if (!this.doc) {
      this.doc = emptyDocument();
      await saveCode(this.doc);
    }
    await setTrusted((await this.currentKey())!, true);
    this.trusted = true;
    this.log("workbook trusted");
    this.changed();
    await this.reload();
    await this.afterTrustChange();
  }

  async revokeTrust(): Promise<void> {
    const key = await this.currentKey();
    if (key) await setTrusted(key, false);
    this.trusted = false;
    this.kernel.restart("trust revoked");
    this.log("trust revoked");
    this.changed();
    await this.afterTrustChange();
  }

  private async afterTrustChange(): Promise<void> {
    try {
      await Office.addin.setStartupBehavior(this.trusted ? Office.StartupBehavior.load : Office.StartupBehavior.none);
    } catch {
      // SharedRuntime 1.1 not available; events then start when the pane opens.
    }
    await Excel.run(async (context) => {
      context.workbook.application.calculate(Excel.CalculationType.full);
      await context.sync();
    });
  }

  // ------------------------------------------------------------------ code

  async saveModules(modules: Module[]): Promise<void> {
    if (!this.doc || !this.trusted) throw new Error("trust the workbook before editing its code");
    const oldKey = await this.currentKey();
    this.doc = { ...this.doc, modules };
    await saveCode(this.doc);
    // The editor's own changes stay trusted; edits made elsewhere change the key and are not.
    if (oldKey) await setTrusted(oldKey, false);
    await setTrusted((await this.currentKey())!, true);
    this.log(`saved ${modules.length} module(s) to the workbook`);
    await this.reload();
    await Excel.run(async (context) => {
      context.workbook.application.calculate(Excel.CalculationType.full);
      await context.sync();
    });
  }

  async reload(): Promise<void> {
    if (!this.trusted || !this.doc) return;
    try {
      const catalog = await this.kernel.loadModules(this.doc.modules);
      for (const e of catalog.errors) this.log(`module ${e.module}: ${e.message}${e.trace ? `\n${e.trace}` : ""}`);
      this.changed();
      if (!this.workbookOpenSent && catalog.events.includes("workbook_open")) {
        this.workbookOpenSent = true;
        void this.dispatch({ type: "workbook_open" });
      }
    } catch (error) {
      this.log(`module load failed: ${message(error)}`);
    }
  }

  // ------------------------------------------------------------------ functions and commands

  async evaluate(item: CallItem): Promise<unknown[][]> {
    if (!this.trusted) {
      throw new CustomFunctions.Error(
        CustomFunctions.ErrorCode.notAvailable,
        "xlfly: code in this workbook is not trusted. Open the xlfly pane to review and trust it.",
      );
    }
    const result = await this.batcher.enqueue(item).catch((error) => {
      throw new CustomFunctions.Error(CustomFunctions.ErrorCode.notAvailable, `xlfly: ${message(error)}`);
    });
    if (!result.ok) throw toFunctionError(result.error.code, result.error.message);
    return wireToMatrix(result.value as Wire).map((row) =>
      row.map((cell) => (typeof cell === "object" ? toFunctionError(cell.error.code, cell.error.message) : cell)),
    );
  }

  async runCommand(name: string): Promise<void> {
    this.log(`running ${name}…`);
    try {
      const result = await this.kernel.runCommand(name);
      if (result.ok) this.log(`${name} finished${result.value != null ? `: ${JSON.stringify(result.value)}` : ""}`);
      else this.log(`${name} failed: ${result.error?.message}${result.error?.trace ? `\n${result.error.trace}` : ""}`);
    } catch (error) {
      this.log(`${name} failed: ${message(error)}`);
    }
  }

  // ------------------------------------------------------------------ events

  private async registerEvents(): Promise<void> {
    try {
      await Excel.run(async (context) => {
        context.workbook.worksheets.onChanged.add(async (event) => {
          if (this.tracker.isOwnWrite(event.address)) return;
          await this.dispatch({
            type: "sheet_change",
            address: event.address,
            worksheetId: event.worksheetId,
            changeType: event.changeType,
            source: event.source,
          });
        });
        context.workbook.worksheets.onSelectionChanged.add(async (event) => {
          await this.dispatch({ type: "selection_change", address: event.address, worksheetId: event.worksheetId });
        });
        await context.sync();
      });
    } catch (error) {
      this.log(`event registration failed: ${message(error)}`);
    }
  }

  private async dispatch(event: Record<string, unknown>): Promise<void> {
    if (!this.trusted || !this.kernel.catalog.events.includes(String(event.type))) return;
    try {
      const result = await this.kernel.dispatchEvent(event);
      for (const e of result.errors as { handler: string; message: string }[]) this.log(`${event.type} handler ${e.handler}: ${e.message}`);
    } catch (error) {
      this.log(`${event.type} failed: ${message(error)}`);
    }
  }

  private changed(): void {
    for (const listener of this.listeners) listener();
  }
}

function message(error: unknown): string {
  return error instanceof Error ? error.message : String(error);
}

function toFunctionError(code: string, text: string): CustomFunctions.Error {
  const known = Object.values(CustomFunctions.ErrorCode) as string[];
  const errorCode = (known.includes(code) ? code : CustomFunctions.ErrorCode.invalidValue) as CustomFunctions.ErrorCode;
  return new CustomFunctions.Error(errorCode, text ? text.slice(0, 255) : undefined);
}

function repeatingArgs(args: unknown): unknown[] {
  if (!Array.isArray(args)) return [];
  return args.filter((a) => a !== undefined && a !== null).map(argToWire);
}

// ------------------------------------------------------------------ bootstrap

const scope = globalThis as any;

if (!scope.__xlfly) {
  const app = new App();
  scope.__xlfly = app;

  CustomFunctions.associate("VERSION", () => `xlfly ${XLFLY_VERSION}`);
  CustomFunctions.associate("PY", (code: string, args: unknown) =>
    app.evaluate({ kind: "eval", code, args: repeatingArgs(args) }),
  );
  CustomFunctions.associate("CALL", (name: string, args: unknown) =>
    app.evaluate({ kind: "call", name, args: repeatingArgs(args) }),
  );

  Office.onReady(async (info) => {
    if (info.host !== Office.HostType.Excel) return;
    await app.init();
    const root = document.getElementById("xlfly-root");
    if (root) bindTaskpane(app, root);
  });
}
