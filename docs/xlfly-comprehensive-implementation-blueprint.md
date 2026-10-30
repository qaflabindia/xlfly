# xlfly — Implementation Blueprint (Revision 3)

**Office Add-in (Office.js) · Excel for Mac, Windows and web · Pyodide (CPython on WebAssembly) in a Web Worker · Two-way interoperability**

Statements are marked **[F]** fact, **[I]** inference, **[U]** unverified (to be settled in Phase 0) where it matters.

## Revision history

| Area | Revision 2 | Revision 3 |
|---|---|---|
| Platform | Windows only | Mac first; same build runs on Windows (Microsoft 365) and the web |
| Host | Excel-DNA XLL, C#, .NET Framework 4.8.x | Office Add-in, TypeScript, shared runtime |
| Python | Embeddable CPython in a child process | Pyodide (CPython 3.14 on WebAssembly) in a Web Worker |
| Isolation | Child process in a Job Object | Web Worker; terminated and restarted by the host |
| IPC | Anonymous pipes, binary framing | `postMessage`, symmetric JSON RPC |
| Python functions in cells | Dynamically registered Excel functions | `=XLFLY.CALL("NAME", …)` and `=XLFLY.PY(code, …)` (static function metadata) |
| Charts | Python-rendered images only | Native, editable Excel charts **and** matplotlib images |
| Installation | Signed installer, XLL trust | Manifest deployment; no native code on the machine |

Reason for the change: XLL add-ins and Excel-DNA are Windows-only; Excel for Mac has no COM and does not load XLLs [F]. The Office Add-in platform is the supported cross-platform extension model for Excel on Mac, Windows and the web [F].

Revision 2's execution-context rules, trust model and failure analysis carry over unchanged in substance.

## 1. Product

### 1.1 Definition

xlfly is an Excel add-in that lets users write Python that runs on their own computer, called from cells, from task-pane commands and from workbook events, with workbook read/write and charting. No paid runtime dependency.

### 1.2 Position against existing options

- **Microsoft Python in Excel (`=PY`)** runs in a Microsoft cloud container with a curated package set [F]. xlfly runs locally and adds commands, events and workbook writes.
- **xlwings Lite** uses the same architecture (Office.js + Pyodide) [F, per vendor]. xlfly is therefore not architecturally unique; differentiation must come from the product layer — governance and trust model, analytics and charting library, enterprise deployment, support. **This is the open strategic question (§13).**
- **PyXLL / Excel-DNA**: Windows-only; not applicable on Mac.

### 1.3 Supported environments

- Excel for Mac, Microsoft 365 (primary target).
- Excel for Windows, Microsoft 365 (Edge WebView2).
- Excel on the web.
- Not supported: Office versions whose add-in runtime is Internet Explorer (no WebAssembly) [F], and any host without the SharedRuntime 1.1 requirement set.

### 1.4 Non-goals

Access to local files, local databases or arbitrary network hosts from Python (browser sandbox) [F]; installing packages from cells; editing Python images with Excel chart tools.

## 2. Architecture

```text
Excel (Mac / Windows / web)
 └─ xlfly add-in — shared runtime page (taskpane.html + app.js)
      Custom functions: XLFLY.PY, XLFLY.CALL, XLFLY.VERSION
      Call batcher · Context guard · Trust manager · Code store
      Office.js services: range read/write, Excel functions, charts, images
      Event bridge · Task pane UI
        │  postMessage, symmetric RPC (protocol v1)
        ▼
    kernel.js — module Web Worker
      Pyodide (CPython 3.14, WebAssembly) · xlfly Python package
      User modules · numpy / pandas / matplotlib (self-hosted, loaded on import)
```

- **Shared runtime** [F]: custom functions, events and the task pane run in one JavaScript runtime, so trust state, the kernel and the code store are shared.
- **Kernel isolation**: Python runs in a dedicated module Web Worker. A hang or crash is contained; the host terminates and restarts the worker and reloads modules. Classic workers are not supported by the current Pyodide release [F, observed].
- **Self-hosted Python**: the Pyodide runtime and packages are served from the add-in's own origin. `npm run fetch:pyodide` downloads packages and verifies each against the SHA-256 in Pyodide's lock file. No third-party CDN at run time.

## 3. Execution contexts

| Context | Entered by | Python may | Python may not |
|---|---|---|---|
| Function | `=XLFLY.PY`, `=XLFLY.CALL` | Compute from arguments, return a value or array | Any call back into Excel |
| Command | Task pane command, event handler | Read/write ranges, call Excel functions, insert charts and images | Run concurrently with another command |

Enforced twice:

1. Python: `xlfly` API calls raise `ContextError` outside command context.
2. Host: every kernel→host request must carry a live command token; tokens exist only while a command or event handler runs. Requests without one are refused with `CONTEXT`.

Rules:

- Worksheet functions are synchronous; returning a coroutine is an error.
- Calls made in the same tick are sent as one batch.
- A worksheet function exceeding 30 s restarts the kernel (synchronous Python cannot be interrupted without cross-origin isolation [F]).
- Commands and event handlers are serialised so that workbook writes never interleave.
- Changes caused by xlfly's own writes are not re-dispatched as `sheet_change` events (address and time-window match) [I: heuristic].

## 4. Python API

```python
import xlfly

@xlfly.func                # =XLFLY.CALL("SCALE", A1:A10, 2)
def scale(values, k): ...

@xlfly.command             # task pane button
async def report():
    rows = await xlfly.Range("Data!A1:B20").read()
    df   = await xlfly.Range("Data!A1:B20").to_frame()
    await xlfly.Range("Out!A1").write(df)
    total = await xlfly.call("SUM", 1, 2, 3)
    await xlfly.chart("column", "Out!A1:B10", "Out!D2", title="Totals")   # native chart
    await xlfly.image(figure, "Out!D20", name="trend")                     # matplotlib PNG
    await xlfly.log("done")

@xlfly.on("sheet_change")
async def changed(event): ...
```

Type mapping (one definition, both directions): number ↔ `float`/`int`; text ↔ `str`; boolean ↔ `bool`; empty ↔ `None`; error ↔ `xlfly.XLError`; dates are serial numbers with `to_excel_date`/`from_excel_date`; 1-D results become a column; DataFrames include a header row; NaN becomes `#NUM!`.

## 5. Protocol

Messages over `postMessage`:

```text
{ v: 1, kind: "req",  id, method, params, context?, deadline? }
{ v: 1, kind: "res",  id, ok: true,  result }
{ v: 1, kind: "res",  id, ok: false, error: { code, message } }
{ v: 1, kind: "note", method, params }
```

Host → kernel: `start`, `load_modules`, `call_functions`, `run_command`, `event`.
Kernel → host: `range_read`, `range_write`, `excel_call`, `chart_native`, `image_insert`, `log`; notifications `stdout`, `stderr`.

Each side numbers its own requests; requests carry deadlines; a version mismatch is refused; closing the channel rejects everything outstanding. Python and JavaScript exchange JSON strings only, so no object proxies cross the boundary.

## 6. Code storage and trust

- Modules live in a workbook custom XML part (`urn:xlfly:code:v1`), stored as base64 JSON so arbitrary source survives XML.
- Each workbook gets a random ID at first use.
- **Trust key** = SHA-256 over length-prefixed (document URL, workbook ID, modules sorted by name). Stored per user via `OfficeRuntime.storage`.
- Code runs only when its key is trusted. Untrusted workbooks return `#N/A` with an explanation.
- Editing in the xlfly pane re-trusts the new code. Any change made elsewhere, or a copy at a different location, requires review and trust again.
- When trusted, the add-in asks Office to load at document open, so events work without opening the pane.

## 7. Charts

- **Native charts** through Office.js: line, scatter, column, bar, area, pie, histogram, box. Editable with Excel's own tools.
- **Python images**: matplotlib figures rendered to PNG and inserted as pictures; a named picture is replaced in place.
- Chart lifecycle work from Revision 2 (definitions, defined-name bindings, duplicate-ID handling, transactional replacement) moves to Phase 3.

## 8. Security

- Python runs inside the browser sandbox of the Office webview: no file system, no process creation, network limited by cross-origin rules [F].
- Kernel→workbook access is limited to command context and the active workbook.
- No native code is installed; there is no XLL, Mark-of-the-Web or antivirus exposure from xlfly binaries [I].
- Python runtime and packages are self-hosted with checksum verification.
- Code in workbooks is untrusted until the user trusts it (§6).
- Residual risk [I]: a trusted command can overwrite workbook data, so trust prompts must show the code.

## 9. Repository layout

```text
addin/
  manifest.xml
  build.mjs  scripts/
  src/app.ts  src/host/  src/kernel/  src/shared/  src/taskpane/  src/functions/
  python/xlfly/  python/tests/
  tests/
docs/
```

## 10. Verification status (this revision)

Verified in the build environment (Linux, Node 22, headless Chromium):

- TypeScript typecheck; 19 unit tests (RPC, value conversion, code store and trust key, kernel client context guard, stale-token refusal, timeout restart with module reload, command serialisation, batching).
- 14 Python unit tests under CPython 3.13.
- Real Pyodide kernel driven through the host client: functions, context guard, a command that reads, calls an Excel function, writes and charts, stdout, events.
- Built bundles in Chromium: kernel starts (Python 3.14.2) in about 4.3 s cold; a 1,000-call batch completes in 24 ms; the task pane against a stubbed Office.js runs trust → save → `=XLFLY.PY` → `=XLFLY.CALL` → error mapping.

### 10.1 Excel for Mac — first real-Excel run (2026-10-30)

Environment: macOS 26.6.2 on Apple Silicon; Microsoft Excel 16.114.3 (build 16.114.26102720, Microsoft 365, universal binary); Node 25.5.0; CPython 3.13.9 for the Python unit tests.

Dependencies: all nine npm packages resolve to the current registry release (`pyodide` 314.0.7, `typescript` 7.0.2, `esbuild` 0.28.2, `@types/office-js` 1.0.615, `@types/office-runtime` 1.0.40, `@types/custom-functions-runtime` 1.6.13, `@types/node` 26.6.6, `office-addin-dev-certs` 3.0.1, `office-addin-manifest` 3.0.1). `npm audit`: 0 vulnerabilities across 65 packages. `npm run check` passes on macOS (typecheck, 19 JS tests, 14 Python tests, Pyodide kernel integration, build). `npm run fetch:pyodide` fetched 13 wheels (numpy, pandas, matplotlib and dependencies), each matching its lock-file SHA-256.

Verified in Excel for Mac [F, observed]:

| Phase 0 item | Result |
|---|---|
| Sideload via the `wef` folder | Pass. The folder did not exist before the first sideload; `sideload-mac.sh` creates it. Excel must be quit and reopened. |
| Shared runtime, module worker, WebAssembly in WebKit | Pass. The **Python** button appears on the Home tab; the kernel starts and the task pane loads modules. |
| `=XLFLY.CALL` with a range and a spilled result | Pass. `GROWTH` over `A1:A10` spilled into `D12:D21`. |
| Trust flow | Pass for **Trust this workbook**; functions return results only after trusting. |
| Command: read, write, native chart | Pass. `SUMMARY` read `A1:A10`, wrote `C1:C2` and inserted an editable column chart at `E2`. |
| Kernel restart from the pane | Pass. **Restart** recovered; the log shows the worker closing and reloading modules. |
| `=XLFLY.CALL` on a command name | Returns `#VALUE!` instead of an error that says the name is a command. Usability gap; see Phase 1. |

Setup findings:

- `office-addin-dev-certs install` needs the user's macOS password; it cannot run unattended.
- The dev certificate is trusted through the macOS keychain. Tools that use their own CA bundle (for example Anaconda's `curl`) reject `https://localhost:3000` even though Excel accepts it; `/usr/bin/curl` succeeds.
- Each command run inserts a new chart; the previous one is not replaced (Excel names them Chart 1, Chart 2, …). Chart replacement is Phase 2 work.
- `npm run validate` fails with "Manifest Version Too Low" because the manifest version is `0.1.0.0`. This blocks AppSource only, not sideloading; raise it to `1.0.0.0` before release.

**Not yet verified [U]:** `=XLFLY.PY` inline evaluation in Excel; matplotlib images; `sheet_change` and `selection_change` events and self-write suppression; save, close and reopen with trust preserved; trust revocation after an outside edit; hang → automatic 30 s restart; numpy/pandas/matplotlib cold-load time inside Excel; 10,000-cell recalculation time; undo history after writes; Excel on Windows and on the web.

## 11. Delivery plan

### Phase 0 — Real-Excel validation (1–2 weeks)

On Excel for Mac (current channel), Windows (Microsoft 365) and the web:

1. Sideload; shared runtime loads; module worker and WebAssembly run in the Excel webview.
2. `=XLFLY.PY` / `=XLFLY.CALL` with ranges, spills and errors; 10,000-cell recalculation time.
3. Trust flow; save, close, reopen; trust revoked when code is edited outside the pane.
4. Command: read, write, native chart, matplotlib image; events with self-write suppression.
5. Kernel hang → restart; Excel stays responsive.
6. numpy/pandas/matplotlib cold-load time from self-hosted packages.
7. Undo behaviour after writes.

Exit: each item passes or has an accepted mitigation.

Status on Excel for Mac (§10.1): items 1 and 2 (`XLFLY.CALL` only) pass; item 3 passes for trusting a workbook; item 4 passes for read, write and native chart; the rest are open. Windows and the web are not started.

### Phase 1 — Hardening (3–4 weeks)

Clear error when `XLFLY.CALL` names a command rather than a function; a way to run a command without the task pane (ribbon button or keyboard shortcut); progress indicator for kernel start; persistent kernel across task pane open/close; better editor (syntax highlighting); package selection per workbook; structured error display with tracebacks; telemetry-free diagnostics export.

### Phase 2 — Chart lifecycle (4–6 weeks)

Chart definitions in the workbook, defined-name source bindings, refresh, duplicate detection, transactional replacement.

### Phase 3 — Differentiating library (scope depends on §13)

### Phase 4 — Release (3–4 weeks)

Production hosting, admin-center deployment guide, AppSource submission if commercial, compatibility matrix.

Planning total after Phase 0: approximately 13–18 weeks plus Phase 3. Estimates, not measurements.

## 12. Failure modes

| Failure | Response |
|---|---|
| Kernel hang in a worksheet function | 30 s timeout, worker terminated and restarted, modules reloaded |
| Kernel crash | Worker error → status `failed`; next call restarts |
| Function tries to use the workbook | `ContextError` in Python; `CONTEXT` refusal at host |
| Untrusted or modified code | Not executed; `#N/A` with explanation |
| Protocol mismatch | Refused |
| Large ranges | Reads and writes above 2,000,000 cells refused |
| Package fails to load | Logged to the pane; import error reported in the cell or command |
| Event loop from own writes | Suppressed by write tracker |
| Unsupported Office version | Manifest requirement set prevents installation |

## 13. Open decisions

1. **Differentiator.** Given that xlwings Lite uses the same architecture, what makes xlfly worth choosing: governance and trust, a statistics and charting library, enterprise deployment, or something else?
2. Internal use or commercial distribution (affects AppSource and hosting)?
3. Do users need local files or databases? If yes, a local companion service is required, which the browser sandbox otherwise forbids.
4. Which packages beyond numpy, pandas and matplotlib?

## 14. References

- Office Add-ins platform overview — https://learn.microsoft.com/office/dev/add-ins/overview/office-add-ins
- Excel custom functions — https://learn.microsoft.com/office/dev/add-ins/excel/custom-functions-overview
- Shared runtime — https://learn.microsoft.com/office/dev/add-ins/develop/configure-your-add-in-to-use-a-shared-runtime
- Sideload on Mac — https://learn.microsoft.com/office/dev/add-ins/testing/sideload-an-office-add-in-on-mac
- Pyodide — https://pyodide.org
- Microsoft PY function — https://support.microsoft.com/en-US/Excel/python/py-function
- xlwings Lite vs Python in Excel — https://www.xlwings.org/lite/vs-python-in-excel
