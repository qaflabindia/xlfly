# xlfly add-in

Local Python for Excel, as an Office Add-in. One code base for Excel on Mac, Windows (Microsoft 365) and the web. Python runs on the user's computer in Pyodide (CPython compiled to WebAssembly) inside a Web Worker; workbook data does not leave the machine.

## What it does

- `=XLFLY.PY(code, args...)` evaluates Python in a cell. Arguments are `arg1`, `arg2`, … and `args`; the last expression is the result. Arrays spill.
- `=XLFLY.CALL("NAME", args...)` calls a function registered with `@xlfly.func` in the workbook's modules.
- `@xlfly.command` functions run from the task pane and can read and write the workbook, call Excel functions, and insert native charts or matplotlib images.
- `@xlfly.on("sheet_change" | "selection_change" | "workbook_open")` handlers react to workbook events.
- Python modules are stored inside the workbook (custom XML part) and run only after the user trusts that workbook on that computer.

```python
import xlfly

@xlfly.func
def growth(values, rate):
    """=XLFLY.CALL("GROWTH", A1:A10, 0.05)"""
    return [[v * (1 + rate)] for row in values for v in row]

@xlfly.command
async def report():
    df = await xlfly.Range("Data!A1:C100").to_frame()          # needs pandas
    await xlfly.Range("Report!A1").write(df.describe().reset_index())
    await xlfly.chart("column", "Report!A1:D9", "Report!F2", title="Summary")
```

Worksheet functions are synchronous and cannot touch the workbook (Excel forbids side effects from functions, and a callback would deadlock). Commands and event handlers are `async` and use `await`.

## Develop on a Mac

Requirements: Node.js 22+, Excel for Mac (Microsoft 365).

One step:

```bash
npm run setup:mac   # install, fetch Python packages, certificate, build, sideload
npm run serve       # keep running; then restart Excel
```

Or step by step:

```bash
npm install
npm run fetch:pyodide   # optional: self-host numpy, pandas, matplotlib (checksums verified)
npm run certs           # creates and trusts a localhost HTTPS certificate (asks for your password)
npm start               # builds dist/ and serves it at https://localhost:3000
npm run sideload:mac    # copies the manifest into Excel's add-in folder
```

Restart Excel, then **Home → Add-ins → xlfly** (or the **Python** button on the Home tab).

Windows: sideload `dist/manifest.xml` through a shared-folder trusted catalog, or with Microsoft's `office-addin-debugging` tool. Web: **Home → Add-ins → More Add-ins → Upload My Add-in**.

## Checks

```bash
npm run typecheck     # TypeScript
npm test              # RPC, value conversion, code store, trust key, kernel client, batching
npm run test:py       # Python package under CPython
npm run test:kernel   # real Pyodide kernel driven through KernelClient
npm run check         # all of the above plus a build
```

## Deploy

```bash
XLFLY_BASE_URL=https://addins.example.com/xlfly npm run build
```

Host `dist/` on any HTTPS static host, then deploy `dist/manifest.xml` to users through the Microsoft 365 admin center (Integrated apps). Set `XLFLY_PACKAGE_BASE_URL` only if Python packages are hosted somewhere other than `dist/pyodide/`.

## Layout

```text
manifest.xml               add-in manifest ({{BASE_URL}} filled at build)
src/app.ts                 shared runtime: custom functions, trust, events, task pane wiring
src/host/                  kernel client, batching, Office.js services, workbook storage
src/kernel/                Web Worker: Pyodide + RPC
src/shared/                RPC protocol, value conversion, code-part format
src/taskpane/              task pane UI
src/functions/functions.json  custom function metadata
python/xlfly/              Python API and kernel runtime
tests/, python/tests/      tests
```

## Known limits

- Excel custom-function names are declared statically in `functions.json`, so Python functions are called through `XLFLY.CALL("NAME", …)` rather than appearing as their own Excel functions.
- A worksheet function that runs longer than 30 s restarts the kernel; synchronous Python cannot be interrupted any other way without cross-origin isolation.
- Pyodide runs in the browser sandbox: no access to local files, local databases or arbitrary network hosts (cross-origin rules apply).
- Cold start is about 4 s in Chromium before packages load; numpy/pandas/matplotlib add download and import time on first use.
- Not yet verified inside real Excel. See the blueprint's Phase 0 list.
