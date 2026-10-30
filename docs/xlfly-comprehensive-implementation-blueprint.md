# xlfly — Implementation Blueprint (Revision 2)

**Local Python runtime for Excel · Excel-DNA host (C#, .NET Framework 4.8.x) · Out-of-process CPython kernel · Two-way interoperability · Python charting**

## Revision summary

Revision 1 proposed a hand-written C++20 XLL driving a standard-library-only Python charting worker, with execution of user Python explicitly out of scope. Revision 2 changes the product and the architecture:

| Area | Revision 1 | Revision 2 |
|---|---|---|
| Product | Charting add-in | Local Python runtime for Excel; charting is one capability |
| User Python | Non-goal | Core capability, under a macro-grade trust model |
| Host | Hand-written C++20 XLL | Excel-DNA (forked, pinned), C#, .NET Framework 4.8.x |
| Python | Standard library only | Embeddable CPython + numpy + pandas (pinned) |
| Interop | Host → worker only | Symmetric RPC: Excel ↔ Python |
| IPC | Named pipes | Anonymous pipes inherited at process creation |
| Worker containment | Unspecified | Windows Job Object, kill-on-close |
| Source binding | Fixed `$A$1` addresses | Hidden workbook-level defined names |
| Chart output | SVG only | SVG, PNG fallback decided by spike |
| Paid dependencies | — | None |
| Plan | M0 build ceremony first | Phase 0 show-stopper sweep and spike first |

Statements are marked where relevant as **[F]** fact, **[I]** inference, **[U]** unverified and to be confirmed in Phase 0.

## 1. Product definition

### 1.1 Positioning

xlfly lets Excel users write Python that runs **locally**, reads and writes the workbook, and produces charts, without data leaving the machine and without a paid third-party runtime.

Existing options and where xlfly differs:

- **Microsoft Python in Excel (`=PY`)** runs in a Microsoft cloud container with a curated package set; code cannot reach local files, databases or the network, and returns values only to its own cell [F, per Microsoft documentation and Q&A as of 2024; confirm current terms in Phase 0]. xlfly targets the cases it excludes: local or regulated data, offline use, custom packages, local system access, and workbook automation.
- **PyXLL** is a mature in-process runtime but a paid, closed-source dependency. Rejected on vendor-dependency grounds.
- **xlwings** (open-source core) drives Excel through COM and VBA for UDFs; weaker performance and robustness for a product runtime [I]. xlwings Lite runs Pyodide in the browser and cannot load arbitrary native packages [I].

Code written for xlfly should port to and from Microsoft `=PY` with minimal change: the same range-to-DataFrame conventions and figure-as-image output.

### 1.2 Target user (to be confirmed — see §16)

Windows Excel users in organisations that require local execution, custom Python packages, or Python-driven workbook automation.

### 1.3 First production release capabilities

1. Python functions callable from cells (`@xlfly.func`).
2. Python commands and event handlers that read and write the workbook (`@xlfly.command`, `@xlfly.on`).
3. Python-rendered charts inserted, edited and refreshed in the workbook.
4. Code stored in the workbook, optionally linked to external `.py` files.
5. Save, close and reopen preserving code, chart definitions and outputs.
6. Code execution gated by a trust model equivalent to or stricter than VBA macros.

### 1.4 Non-goals

- Mac, Excel on the web, and 32-bit Excel in the first release. Workbooks opened there show the last stored outputs and cannot recalculate xlfly functions.
- Editing Python-rendered charts through Excel's native chart tools.
- Installing packages from inside a cell or at run time.
- Undocumented Excel hooks or reverse engineering.

## 2. Architecture

```text
Excel (64-bit, Microsoft 365, Windows)
 └─ xlfly.xll — Excel-DNA host, C#, .NET Framework 4.8.x
      Function registry · Command router · Event bridge
      Main-thread dispatcher (QueueAsMacro) · Context guard
      Trust manager · Code store · Chart artifact manager
      Kernel supervisor · RPC endpoint
        │  anonymous pipes (inherited handles), Job Object
        ▼
    xlfly kernel — embeddable CPython 3.x (pinned) + numpy + pandas
      RPC endpoint · Module loader · Function/command registry
      Type converters · User code sandbox boundary
      Chart engine (geometry · layout · SVG/PNG)
```

Division of responsibility:

- **Host (C#):** everything that touches Excel — registration, COM and C API access, main-thread scheduling, Ribbon and task pane, trust decisions, workbook persistence, chart insertion.
- **Kernel (Python):** everything users write and everything analytical — user modules, transforms, statistics, rendering. The kernel never touches Excel directly; all workbook access is an RPC to the host.

Excel and the kernel are separate failure domains. A kernel crash, hang or memory blow-up must not terminate Excel; Excel exiting must terminate the kernel.

### 2.1 Why Excel-DNA

Excel-DNA packages .NET code as a standard `.xll` and already provides: function registration and runtime re-registration, memory and exception marshalling at the Excel boundary, asynchronous and streaming functions, main-thread scheduling (`ExcelAsyncUtil.QueueAsMacro`), COM object model access, Ribbon and custom task panes from the same add-in, and single-file packing [F, per Excel-DNA documentation].

Licence: zlib, permissive, commercial use without fee [F]. Maintenance is active (1.9.0 release candidates in 2025; 1.10.0 Native AOT preview listed July 2026) [F, per NuGet listings; stable 2026 release unconfirmed].

Vendor risk control: fork Excel-DNA into the organisation, pin to a reviewed commit, build it from source in CI. The project then has no external party able to block it.

### 2.2 Why .NET Framework 4.8.x rather than modern .NET

- Only one modern .NET runtime can load into a process; an add-in built for one modern .NET major version cannot coexist in the same Excel instance with one built for another. .NET Framework add-ins coexist with whichever modern runtime is loaded [F, per Excel-DNA maintainers].
- Modern .NET requires the Desktop Runtime to be installed; .NET Framework 4.8.x ships with Windows [F].
- .NET 8 and 9 reach end of support on 10 November 2026; .NET 10 LTS runs to November 2028 [F]. Modern .NET would impose a forced runtime migration every 2–3 years.
- The host is thin; computation runs in Python. Modern .NET performance gains do not matter here [I].

Revisit only for managed enterprise estates where the runtime and co-installed add-ins are controlled.

### 2.3 Why an out-of-process kernel

- User code will hang, leak and crash in native extensions; a separate process can be killed and cannot take Excel down [I].
- Cancellation is reliable (terminate the process); in-process Python threads cannot be safely aborted [F].
- Excel's multithreaded recalculation and the Python GIL do not interact.

Cost: per-call IPC latency. Mitigated by batching all xlfly function calls within one recalculation into a single round trip (§4.3). Latency budget to be measured in Phase 0.

### 2.4 Rejected alternatives

| Alternative | Reason rejected |
|---|---|
| Hand-written C++ XLL | Highest effort and risk; re-implements what Excel-DNA provides |
| PyXLL | Paid, closed-source dependency |
| xlwings classic | VBA/COM-based UDF path; product robustness and performance [I] |
| Microsoft `=PY` only | Cloud execution; no local data, custom packages or workbook writes |
| Office.js add-in | No local CPython without a local server; different product |
| C#-only (no Python) | Fails the requirement that users write Python |

## 3. Execution contexts and threading

### 3.1 Excel constraints [F]

- A worksheet function may only return a value to its calling cell. It cannot write other cells, insert shapes or alter the workbook.
- Workbook changes are permitted only in macro context: commands, Ribbon actions, or work scheduled through `QueueAsMacro`.
- COM and C API calls must run on Excel's main thread.
- Workbook changes made by code clear Excel's undo stack.

### 3.2 xlfly contexts

| Context | Entered by | Python may | Python may not |
|---|---|---|---|
| **Function** | Cell recalculation calling `@xlfly.func` | Compute from arguments; return a value or array | Any call back into Excel |
| **Command** | Ribbon, `@xlfly.command`, event handler | Read and write ranges, call Excel functions, insert charts | Write to workbooks other than the originating one (default) |

**Deadlock rule.** In function context Excel's calculation waits on Python. If Python requested main-thread work, both sides would wait forever. The host therefore rejects every kernel→host request carrying a function-context token, with a typed error. All function inputs must arrive as arguments, which also keeps Excel's dependency tracking correct. Read-only access to other ranges from functions is a possible later extension and must be documented as untracked by Excel.

In command context, Python awaits each host request; the host executes it through `QueueAsMacro` when Excel is ready, keeping Excel responsive between calls.

### 3.3 Dispatch rules

- Separate kernel queues for function work and command work, so a long command never stalls recalculation and a write-triggered recalculation cannot deadlock against the running command.
- Every RPC has a deadline. Excel busy (edit mode, modal dialog) delays queued work; Python sees a pending call, not an error, until the deadline.
- xlfly suppresses its own change events while it writes, preventing event storms.
- Workbook close cancels all jobs bound to that workbook and rejects their remaining callbacks.

## 4. Two-way interoperability

### 4.1 Excel → Python

1. **Functions.** `@xlfly.func` registers a Python function as an Excel function. The host uses Excel-DNA runtime registration; module reload re-registers without restarting Excel.
2. **Commands.** `@xlfly.command` exposes entry points to the Ribbon and to macros.
3. **Events.** The host forwards workbook open/close, sheet change, calculation completed and selection events as notifications. `@xlfly.on("sheet_change")` subscribes.

### 4.2 Python → Excel (command context only)

```python
import xlfly

@xlfly.command
def build_report():
    df = xlfly.Range("Data!A1:D10000").to_frame(header=True)   # bulk read
    out = summarise(df)
    xlfly.Range("Report!A1").from_frame(out)                   # bulk write
    xlfly.call("XLOOKUP", "key", xlfly.Range("T!A:A"), xlfly.Range("T!B:B"))
    xlfly.chart(make_figure(out), anchor="Report!F2", chart_id="summary")
```

All transfers are bulk; the API offers no per-cell loop over RPC.

### 4.3 Function batching

During one recalculation the host collects pending xlfly function calls, sends them as one batch, and returns results through Excel-DNA's asynchronous function mechanism. Long-running functions return `#BUSY`-style pending state until complete [U: exact Excel presentation to be confirmed].

### 4.4 Type mapping (single definition, both directions)

| Excel | Python |
|---|---|
| Number | `float` |
| Text | `str` |
| Boolean | `bool` |
| Empty | `None` |
| Error (`#N/A`, `#DIV/0!`, …) | `xlfly.XLError` (never silent `NaN`) |
| Date | `float` by default; `datetime` on request or by cell format, honouring 1900/1904 date systems |
| Range | 2-D `list`, `numpy.ndarray` or `pandas.DataFrame` (explicit header option) |
| Spilled result | Dynamic array |

Converters are shared by functions and commands and covered by round-trip property tests.

## 5. RPC protocol

### 5.1 Transport

- Anonymous pipes created by the host and inherited by the kernel at `CreateProcess`. No pipe name exists to squat [I].
- Kernel placed in a Job Object with `JOB_OBJECT_LIMIT_KILL_ON_JOB_CLOSE`, plus memory limits.
- Kernel launched by absolute path from the add-in directory, with a clean environment and Python isolated mode (`-I`), so `PYTHONPATH`, `PYTHONHOME` and `PYTHONSTARTUP` cannot inject code.

### 5.2 Framing

```text
Header (little-endian):
  magic            4 bytes  "XLFY"
  protocol_major   uint16
  protocol_minor   uint16
  kind             uint16   REQUEST | RESPONSE | NOTIFICATION | CANCEL
  flags            uint16
  payload_length   uint32   bounded by negotiated maximum
  message_id       uint64   unique per direction
  context_token    16 bytes job and context (function | command)
Payload:
  UTF-8 JSON control body
  optional typed binary blocks for bulk arrays (dtype, shape, bytes)
```

Model: symmetric RPC in the style of JSON-RPC 2.0 / Language Server Protocol. Either side may initiate requests.

### 5.3 Methods

Host → kernel: `hello`, `load_modules`, `call_functions` (batch), `run_command`, `event`, `render_chart`, `cancel`, `shutdown`.
Kernel → host: `register_functions`, `range_read`, `range_write`, `excel_call`, `chart_insert`, `log`.

Every request carries a deadline and maximum response size. Version negotiation in `hello` precedes any other message; incompatible majors refuse to communicate.

### 5.4 Job state machine

```text
CREATED → VALIDATED → QUEUED → RUNNING → SUCCEEDED
   │          │          │        ├──→ FAILED
   └──────────┴──────────┴────────┴──→ CANCELLED | TIMED_OUT

For chart and write jobs, after SUCCEEDED:
RENDERED → COMMITTING → COMMITTED | ROLLED_BACK | ORPHANED
```

## 6. Code storage and calculation model

### 6.1 Where code lives

- **Embedded:** Python modules stored in a workbook custom XML part, so the workbook is self-contained.
- **Linked:** workbook references external `.py` files for version control. Per-workbook choice, recorded in the part.

### 6.2 Calculation model

Release 1 uses **pure functions**: a cell's result depends only on its arguments; Excel's dependency graph drives recalculation. Module-level state is permitted but documented as invisible to Excel. An ordered notebook mode (as in Microsoft `=PY`) is deferred until pure-function behaviour is stable.

### 6.3 Python environment

- Embeddable CPython distribution, pinned minor version, shipped in the signed install directory.
- numpy and pandas pinned; further packages added per deployment through a reviewed lock file.
- No package installation from cells or at run time.

## 7. Trust model

Workbook-embedded code has macro-level power. Rules:

1. Code execution is **off by default** for every workbook.
2. Code runs only when the workbook is in a trusted location, signed by a trusted publisher, or explicitly trusted by the user after viewing the code summary.
3. Files carrying Mark of the Web are never auto-trusted.
4. Trust decisions are stored per user, keyed to workbook identity and code hash; any code change revokes trust.
5. Commands write only to the originating workbook unless the user grants wider scope.
6. Administrators can disable or restrict xlfly through policy registry keys.

## 8. Charting

### 8.1 Engine contract

```python
@dataclass(frozen=True)
class ChartArtifact:
    chart_id: str
    mime_type: str        # image/svg+xml or image/png
    content: bytes
    width_px: int
    height_px: int
    sha256: str
```

The engine receives validated data and a specification and returns an artifact; it never writes to Excel. Output format: SVG if Phase 0 confirms insertion, rendering fidelity and save/reopen behaviour [U]; otherwise PNG at 2× DPI. Text from cells is XML-escaped; SVG output additionally passes a sanitiser (no scripts, external resources, foreign objects or links).

### 8.2 Chart definition and identity

```text
ChartDefinition
  chart_id · schema_version · source_bindings · chart_type
  transformations · axes · styles · layout
  refresh_policy · last_render_hash · shape_identity
```

- Stored in a workbook custom XML part.
- Source bindings use **hidden workbook-level defined names**, which Excel adjusts on row/column insertion, deletion and sheet rename [F]. Structured tables bind by table and column name.
- Workbook and sheet IDs are minted by xlfly and stored in the custom XML part; Excel provides no stable UUIDs [F].
- Shapes carry the `chart_id` in documented shape metadata. **Duplicate IDs** (sheet copy, shape copy-paste) are detected on load and on refresh; the copy receives a new ID and a copied definition.
- Missing sources are marked unresolved, never silently rebound.

### 8.3 Fingerprint

`H = SHA-256(len‖schema ‖ len‖values ‖ len‖spec)` with length-prefixed canonical encoding; `-0.0` normalised to `0.0`; NaN canonicalised; Excel errors and dates encoded by type tag; specification serialised with sorted keys.

### 8.4 Transactional replacement

Render → validate artifact and checksum → insert new shape → apply position and size → update definition → remove old shape → commit. A recovery journal in the custom XML part records in-flight replacements; on load, incomplete replacements are reconciled (keep the valid shape, remove the orphan). Excel provides no workbook-wide atomicity [F].

### 8.5 Undo

Insertion and refresh clear Excel's undo stack [F]. Release 1: refresh only on explicit command or opted-in policy; the first operation per session warns the user.

## 9. User interface

- Ribbon tab and custom task pane from the Excel-DNA add-in: code editor/module list, trust status, function list, chart designer, job and log view.
- Chart designer sections: Data, Chart, Axes, Style, Analysis, Output.
- No standalone Win32 dialogs.

## 10. Security model

Trust boundaries:

1. **Workbook → host:** values, names, metadata and embedded code are untrusted until the trust manager admits them.
2. **Host ↔ kernel:** every message validated for size, schema and context; function-context callbacks rejected.
3. **Kernel → workbook:** writes scoped to the originating workbook; chart artifacts validated before insertion.
4. **Installer/update → system:** signed packages, integrity verification.

Distribution constraints:

- Microsoft 365 Excel blocks XLLs carrying Mark of the Web by default since 2023 [F]. Distribution uses managed deployment (Intune/SCCM) or a signed installer placing files in a trusted location.
- XLLs are a known malware delivery format and are scrutinised by endpoint security [F]. All binaries — `xlfly.xll`, the Excel-DNA fork, the Python launcher — are code-signed.

Threat model statement: a malicious process running as the same user can tamper with a per-user install. That is out of scope; per-machine installation is offered where it is in scope.

Dependency policy: Excel-DNA (fork), CPython, numpy, pandas — pinned, licence-reviewed (zlib, PSF, BSD), vulnerability-scanned. No paid or closed-source runtime dependencies.

## 11. Repository layout

```text
host/
  Xlfly.Host/             C# Excel-DNA add-in
  Xlfly.Host.Tests/       unit tests (no Excel)
kernel/
  xlfly/                  public Python API for users
  xlfly_kernel/           RPC server, loader, converters, chart engine
  tests/
protocol/
  schemas/                message schemas, version history
  conformance/            cross-language contract tests
third_party/
  exceldna/               pinned fork (submodule)
tests/excel/              Excel-hosted integration tests
installer/
docs/
```

## 12. Build and CI

- Build: .NET SDK / MSBuild targeting `net48`, Excel-DNA fork built from source, CPython embeddable distribution, pinned wheels.
- CI stages: lint and format → host build and unit tests → kernel tests → protocol conformance (C# ↔ Python) → security (static analysis, protocol fuzzing, dependency audit) → packaging and signing → Excel-hosted integration tests.
- Excel-hosted tests require licensed Excel on an interactive desktop; hosted CI does not provide this [F]. A dedicated, human-maintained Windows machine or self-hosted runner is provisioned in Phase 0.

## 13. Test strategy

- **Unit:** converters, geometry, statistics, protocol framing, trust decisions.
- **Contract:** C# and Python agree on framing, types, versions and errors.
- **Property:** type round-trips; `corr(X, X) = 1` for nonconstant X; histogram bin counts sum to N.
- **Integration (real Excel):** load/unload, registration, function batching, command writes, events, chart insertion, save/reopen.
- **Concurrency:** function-context callback rejected; write-triggered recalculation; workbook closed mid-command; Excel busy.
- **Recovery:** kernel crash, hang and kill; interrupted chart replacement; corrupted custom XML part.
- **Compatibility:** supported Microsoft 365 channels, DPI, locale, co-installed add-ins on modern .NET.
- **Performance:** `T = T_read + T_serialize + T_ipc + T_compute + T_render + T_insert`, plus peak memory and Excel responsiveness. Targets set after Phase 0 baselines.

## 14. Delivery plan

### Phase 0 — Show-stopper sweep and spike (2 weeks)

Spike: Excel-DNA `=XLFLY.PY(code, range)` → out-of-process kernel → DataFrame result; one command writing a range and inserting a chart; save/reopen.

Show-stopper register — each item green or covered by an accepted fallback before Phase 1:

| # | Risk | Test | Mitigation / fallback |
|---|---|---|---|
| 1 | IT policy blocks add-in; Mark of the Web | Install on locked-down image | Signing, managed deployment, trusted location |
| 2 | Endpoint security quarantine | Defender and customer EDR | Signing, vendor submissions, embeddable CPython not PyInstaller |
| 3 | Mac/web required | User inventory | Out of scope, or Microsoft `=PY` for those users |
| 4 | 32-bit Excel present | User inventory | x86 build or explicit exclusion |
| 5 | No admin rights | Install as standard user | Per-user install and registration |
| 6 | Data residency | Compliance review | Determines whether `=PY` is a fallback |
| 7 | Modern .NET add-in conflict | Co-install test | .NET Framework host (by design) |
| 8 | SVG fidelity/persistence | Spike | PNG at 2× DPI |
| 9 | IPC latency | 10,000-cell batched recalculation | Batching, size limits |
| 10 | Deadlock / reentrancy | Function callback, write-triggered recalc | Context guard, separate queues |

Exit: register resolved; architecture confirmed or amended; plan re-estimated.

### Phase 1 — Foundations (3–4 weeks)

Host skeleton, kernel supervisor, Job Object, anonymous-pipe RPC, version negotiation, type converters, `@xlfly.func` with batching, diagnostics.
Exit: functions return correct values in real Excel; kernel crash leaves Excel running; contract tests pass.

### Phase 2 — Two-way interop (4–6 weeks)

Commands, events, Range API, `excel_call`, context guard, code store (embedded and linked), trust manager.
Exit: §3 and §4 behaviours verified in real Excel, including deadlock and cancellation tests; untrusted code never runs.

### Phase 3 — Charting and persistence (4–6 weeks)

Chart engine (scatter first), artifact manager, defined-name bindings, duplicate-ID handling, transactional replacement, recovery journal.
Exit: create, edit, refresh and save/reopen without identity loss or orphaned shapes.

### Phase 4 — Catalogue and UI (3–5 weeks)

Chart types prioritised by the differentiator in §1, task pane, chart designer.
Exit: numerical, geometry and integration tests pass per chart type.

### Phase 5 — Release (4–6 weeks)

Installer, signing, policy keys, documentation, compatibility matrix.
Exit: reproducible, signed release candidate validated on supported Excel.

Planning total: approximately 20–29 weeks after Phase 0. These are estimates, not measurements; re-estimate at the Phase 0 exit.

## 15. Engineering governance

- Each task has a Maker, an independent Checker and acceptance criteria with recorded evidence.
- Agents may implement, test, refactor and open pull requests within assigned paths. Architecture, protocol major versions, new dependencies, trust-model changes and releases require human approval.
- Any task touching Excel behaviour is **Verified** only with evidence from a real Excel run; otherwise it is **Reviewed** or **Unverified**.
- Ledger records task, revision, outcome and artifact hashes; never workbook contents or secrets.

## 16. Open decisions

1. Internal use only, or sold to external customers?
2. Any requirement for Mac or Excel on the web?
3. Are target machines centrally managed and locked down?
4. May workbook data be processed in Microsoft's cloud?
5. Which analyses or charts form the differentiator that orders Phase 4?

## 17. Failure modes

| Failure | Response |
|---|---|
| Kernel crash or hang | Job Object isolation, restart, typed cell errors, command failure report |
| Function-context callback | Rejected with typed error |
| Excel busy or modal | Queued work waits until deadline |
| Workbook closed mid-job | Cancel job, reject callbacks, discard results |
| Untrusted or modified code | Not executed; trust revoked on code change |
| Source deleted or moved | Defined name adjusts; deletion marks binding unresolved |
| Duplicate chart ID | Reassign ID on copy |
| Interrupted chart replacement | Recovery journal reconciliation |
| Protected sheet | Detect and report |
| Protocol mismatch | Refuse communication |
| Add-in blocked by policy or Mark of the Web | Diagnostic message with deployment guidance |
| Event storm from own writes | Suppression during writes |

## 18. Production readiness

- [ ] Add-in loads and unloads reliably on supported builds, alongside modern .NET add-ins.
- [ ] Python functions return correct values; batching meets measured targets.
- [ ] Commands and events read and write correctly; deadlock tests pass.
- [ ] Untrusted code never executes.
- [ ] Kernel failure never terminates Excel; Excel exit always terminates the kernel.
- [ ] Charts survive save/reopen; refresh leaves no duplicate or orphaned shapes.
- [ ] Protocol and workbook schema migrations tested.
- [ ] Installation, repair and uninstallation repeatable without admin rights.
- [ ] All binaries signed and traceable to source revisions.
- [ ] User documentation and diagnostics complete.

## 19. Documentation package

```text
docs/
  00-product-requirements.md
  01-system-architecture.md
  02-adr-host-and-runtime.md
  03-execution-contexts.md
  04-rpc-protocol.md
  05-type-mapping.md
  06-calculation-model.md
  07-trust-model.md
  08-python-api.md
  09-chart-engine.md
  10-workbook-persistence.md
  11-security-model.md
  12-test-strategy.md
  13-build-and-release.md
  14-governance.md
  15-operations-and-recovery.md
```

## 20. References

- [Excel-DNA](https://github.com/Excel-DNA/ExcelDna) · [Excel-DNA licence (zlib)](https://github.com/Excel-DNA/Registration/blob/dc599090c480733046b4c0a266b65502cddfeeb9/LICENSE.txt)
- [Programming with the C API in Excel](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel)
- [Asynchronous User-Defined Functions](https://learn.microsoft.com/en-us/office/client-developer/excel/asynchronous-user-defined-functions)
- [Microsoft PY function](https://support.microsoft.com/en-US/Excel/python/py-function)
- [.NET 8 and .NET 9 end of support](https://devblogs.microsoft.com/dotnet/dotnet-8-9-end-of-support/)
- [Excel blocks internet-sourced XLL add-ins](https://office-watch.com/2023/microsoft-block-excel-xll-add-ins/)

## 21. Directive

Begin with Phase 0. No product code is written until every show-stopper in §14 is green or covered by an accepted fallback.

The first critical transaction is:

```text
Excel cell → Excel-DNA host → Python kernel → result or chart → Excel
```

Once it is reliable in both directions, recoverable and governed by the trust model, further functions and chart types are modules, not architecture changes.
