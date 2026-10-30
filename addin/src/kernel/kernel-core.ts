// The Python kernel: Pyodide plus the xlfly package, driven over RPC.
// Runs inside a Web Worker in Excel, and in-process under Node for tests.

import type { RpcEndpoint } from "../shared/rpc.ts";

export interface KernelConfig {
  indexURL: string;
  packageBaseUrl?: string;
  /** Python package files, path relative to the package root -> source. */
  python: Record<string, string>;
}

export interface KernelInfo {
  python: string;
  pyodide: string;
  xlfly: string;
}

type LoadPyodide = (options: Record<string, unknown>) => Promise<any>;

const PYTHON_ROOT = "/xlfly";

export async function startKernel(rpc: RpcEndpoint, loadPyodide: LoadPyodide, config: KernelConfig): Promise<KernelInfo> {
  const options: Record<string, unknown> = {
    indexURL: config.indexURL,
    stdout: (text: string) => rpc.notify("stdout", { text }),
    stderr: (text: string) => rpc.notify("stderr", { text }),
  };
  if (config.packageBaseUrl) options.packageBaseUrl = config.packageBaseUrl;
  const py = await loadPyodide(options);

  for (const [path, source] of Object.entries(config.python)) {
    const full = `${PYTHON_ROOT}/${path}`;
    py.FS.mkdirTree(full.slice(0, full.lastIndexOf("/")));
    py.FS.writeFile(full, source);
  }
  py.runPython(`import sys\nif ${JSON.stringify(PYTHON_ROOT)} not in sys.path: sys.path.insert(0, ${JSON.stringify(PYTHON_ROOT)})`);

  const host = py.pyimport("xlfly._host");
  host.install_bridge((method: string, paramsJson: string, token: string) =>
    rpc.request(method, JSON.parse(paramsJson), { context: token }).then((r) => JSON.stringify(r ?? null)),
  );
  const runtime = py.pyimport("xlfly._runtime");

  // Load Pyodide packages (numpy, pandas, ...) named in import statements before running code.
  const loadImports = async (code: string) => {
    try {
      await py.loadPackagesFromImports(code);
    } catch (error) {
      rpc.notify("stderr", { text: `package load failed: ${error instanceof Error ? error.message : String(error)}` });
    }
  };

  rpc.on("load_modules", async ({ modules }: { modules: { name: string; source: string }[] }) => {
    for (const m of modules) await loadImports(m.source);
    return JSON.parse(runtime.load_modules(JSON.stringify(modules)));
  });

  rpc.on("call_functions", async ({ batch }: { batch: { kind: string; code?: string }[] }) => {
    for (const item of batch) if (item.kind === "eval" && item.code) await loadImports(item.code);
    return JSON.parse(runtime.call_functions(JSON.stringify(batch)));
  });

  rpc.on("run_command", async ({ name }: { name: string }, meta) => {
    return JSON.parse(await runtime.run_command(name, meta.context ?? ""));
  });

  rpc.on("event", async ({ event }: { event: unknown }, meta) => {
    return JSON.parse(await runtime.dispatch_event(JSON.stringify(event), meta.context ?? ""));
  });

  return {
    python: py.runPython("import sys; sys.version.split()[0]"),
    pyodide: String(py.version),
    xlfly: py.runPython("import xlfly; xlfly.__version__"),
  };
}
