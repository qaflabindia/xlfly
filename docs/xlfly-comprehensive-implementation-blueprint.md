# xlfly — Comprehensive Implementation Blueprint

**Native Excel XLL · C++20 · Excel C API · Python Charting Engine · AAWDF Agentic Development**

## Objective

Build xlfly as an independently engineered, production-grade Excel charting and analytical add-in, using Microsoft’s documented native extension interfaces, a locally executed Python engine and our AI-agent development framework.

This blueprint covers system architecture, implementation contracts, native integration, chart rendering, build infrastructure, agent orchestration, security, testing, deployment and release governance.

**Primary architectural decision:** A native XLL with an **isolated CPython worker**, rather than Python embedded directly in Excel. The XLL controls Excel integration; Python controls analytical computation and rendering.

## 1. Product boundaries and engineering principles

### Supported environment

- Windows 11 x64 initially, with Windows 10 support subject to servicing and security status.
- Microsoft Excel 64-bit desktop, initially Microsoft 365.
- MSVC C++20, CMake and Windows SDK.
- Microsoft Excel XLL SDK and documented Excel C API.
- CPython 3.x, pinned to a tested minor version.
- Local-only execution without a cloud service or third-party charting API.

The chart renderer will be our own implementation. Python’s standard library suffices for the initial SVG engine; optional numerical dependencies require explicit approval.

### Product capabilities

The first production release should support:

1. Select Excel data, configure a chart and insert it into a workbook.
2. Edit an existing xlfly chart without manually recreating it.
3. Refresh charts when source data changes.
4. Save, close and reopen a workbook while preserving chart output and configuration.

Initial chart catalogue: line, scatter, column, bar, histogram, box plot and heatmap. Advanced statistical and ML visualizations follow after core lifecycle stability.

### Non-goals

We will not reverse-engineer Excel’s proprietary calculation engine, hook undocumented Excel functions, execute arbitrary Python expressions from worksheet cells, or claim that SVG images are natively editable Excel charts. Python-rendered charts are edited through xlfly, not Excel’s native chart-series editor.

## 2. System decomposition

```text
Excel Host
  Workbook · Calculation · Selection · Shapes
         |
Native xlfly Host — C++20
  XLL Lifecycle · C API Adapter · COM Adapter
  Command Router · Chart Designer · Artifact Manager
         |
Native Execution Services
  Job Scheduler · Worker Supervisor
  Named-Pipe IPC · Protocol Validator
         |
xlfly Python Engine
  Typed Data Model · Transformations · Statistical Kernel
  Chart Geometry · Layout Engine · SVG Renderer
         |
Workbook Artifacts
  Chart image · Source binding · Registry · Recovery state
```

Excel and the Python worker are separate failure domains. A worker crash must not terminate Excel; an Excel crash must not leave a privileged or indefinitely running worker behind.

## 3. Execution model and thread safety

The Excel C API is host-provided and not an ordinary thread-safe library.

- **Worksheet UDF execution:** Pure calculations returning Excel-compatible values. No workbook mutation, chart insertion or dialogs.
- **XLL command execution:** Explicit user commands operating in permitted Excel contexts.
- **COM workbook operations:** Workbook, worksheet and shape manipulation in a properly initialized and authorized Excel automation context.
- **Python worker execution:** Data transformations, statistics and rendering, with no Excel API access.

Worker completion uses a native queue and an Excel-safe dispatch mechanism. A Windows message alone does not authorize an Excel C API call. Dispatch must respect Excel readiness, reentrancy and the different constraints of C API versus COM operations.

Official references:

- [Programming with the C API in Excel](https://learn.microsoft.com/en-us/office/client-developer/excel/programming-with-the-c-api-in-excel)
- [Multithreading and Memory Management](https://learn.microsoft.com/en-us/office/client-developer/excel/multithreading-and-memory-management)
- [Asynchronous User-Defined Functions](https://learn.microsoft.com/en-us/office/client-developer/excel/asynchronous-user-defined-functions)

Begin with explicit user-triggered commands and controlled completion processing. Automatic event-driven refresh follows later.

## 4. Native host implementation

```text
xlfly_xll
  ├── xlfly_excel_api
  ├── xlfly_excel_com
  ├── xlfly_commands
  ├── xlfly_execution
  ├── xlfly_ipc
  ├── xlfly_artifacts
  └── xlfly_diagnostics
```

### XLL lifecycle

- `xlAutoOpen`: Initialize add-in and register exported functions and commands.
- `xlAutoClose`: Unregister where appropriate and release resources.
- `xlAutoAdd` / `xlAutoRemove`: Optional installation/removal notifications.
- `xlAutoFree12`: Release XLL-owned memory when Excel requests it.
- `xlAddInManagerInfo12`: Optional Add-in Manager metadata.

Do not initialize Python, launch workers or perform substantial blocking work in `DllMain`.

### Function registration

For M0 register:

```excel
=XLFLY.VERSION()
```

Return a stable version string via documented `xlfRegister` registration, with consistent calling conventions, argument types and return ownership. Register chart creation as an explicit command, not a worksheet UDF.

### C++ memory management

Use RAII wrappers distinguishing:

- XLL-allocated memory.
- Excel-allocated results freed through the appropriate Excel mechanism.
- Static or borrowed memory that must not be freed.
- Returned allocations requiring `xlAutoFree12`.

Every pointer crossing the Excel boundary must have a documented owner.

### Error handling

No C++ exception may escape an exported Excel callback. Catch exceptions at the boundary, convert failures to valid Excel errors or controlled command failures, and log diagnostics. Reject oversized datasets before allocation and check dimension and byte-count arithmetic for overflow.

## 5. Python engine architecture

```text
xlfly-engine/
  xlfly/
    worker/
      server.py
      dispatcher.py
      lifecycle.py
    data/
      types.py
      validation.py
      columnar.py
      missing.py
    transforms/
      filter.py
      group.py
      aggregate.py
      sort.py
    statistics/
      descriptive.py
      correlation.py
      regression.py
      distributions.py
    charts/
      base.py
      line.py
      scatter.py
      bar.py
      histogram.py
      boxplot.py
      heatmap.py
    geometry/
      scales.py
      coordinates.py
      ticks.py
      clipping.py
    layout/
      text.py
      axes.py
      legend.py
      margins.py
    rendering/
      svg.py
      styles.py
      fonts.py
      sanitize.py
    persistence/
      specifications.py
      migrations.py
```

### Engine contracts

A chart renderer receives validated data and a chart specification and returns an artifact without writing directly to Excel.

```python
from dataclasses import dataclass
from typing import Protocol

@dataclass(frozen=True)
class ChartArtifact:
    chart_id: str
    mime_type: str
    content: bytes
    width: int
    height: int
    sha256: str

class ChartRenderer(Protocol):
    def render(self, dataset, spec) -> ChartArtifact:
        ...
```

The renderer must be independent of the output host (Excel, PDF or another application).

### Statistical kernel

Separate data transformations from statistical calculations. Histogram binning policies: fixed count, fixed width, Sturges, Scott and Freedman–Diaconis. For regression, specify algorithm and numerical assumptions. QR-based least squares is preferable to unguarded normal equations for ill-conditioned data when an approved numerical library is available.

### Rendering kernel

Implement SVG layout, text measurement, clipping, tick generation, numeric formatting, legend placement and font fallback. For M2 use controlled fonts and deterministic parameters. A documented Windows text API may provide precise text measurement. Sanitize SVG: prohibit scripts, external resources, remote fonts, foreign objects and executable links.

## 6. Data exchange protocol

Use framed binary transport over Windows named pipes.

```text
Header:
  magic             4 bytes
  protocol_major    uint16
  protocol_minor    uint16
  message_type      uint16
  flags             uint16
  payload_length    uint64
  request_id        16 bytes
Payload:
  UTF-8 JSON for control messages
  Typed binary blocks for bulk data
```

Specify integer byte order explicitly; initial Windows-only protocol uses little-endian.

Operations:

- `HELLO`
- `HEALTH_CHECK`
- `RENDER_CHART`
- `CANCEL_JOB`
- `GET_CAPABILITIES`
- `RENDER_RESULT`
- `ERROR`
- `SHUTDOWN`

Every request carries correlation ID, protocol version, deadline and maximum response size. Never accept executable source code through the protocol.

### Job state machine

```text
CREATED → VALIDATED → QUEUED → RUNNING
                                  ├── SUCCEEDED
                                  ├── FAILED
                                  └── CANCELLED
```

Rendering success is distinct from workbook insertion success. A workbook may close, be protected or become busy before insertion. Track artifact commit status separately.

## 7. Excel data acquisition

For a selected range:

1. Identify the Excel application and workbook instance.
2. Resolve the active selection.
3. Validate supported range shape.
4. Read values in bulk.
5. Preserve source metadata and Excel errors.
6. Convert to xlfly’s typed data model.
7. Release Excel-owned resources correctly.

Prefer the Excel C API when it provides the necessary operation; use COM for other workbook operations. Never read live worksheet objects from an arbitrary background thread. Transfer an immutable snapshot to Python.

### Data identity

```json
{
  "source_kind": "excel_range",
  "workbook_id": "wb-uuid",
  "sheet_id": "sheet-uuid",
  "address": "$A$1:$B$100",
  "header_rows": 1,
  "date_system": "1900",
  "binding_mode": "fixed_range"
}
```

For structured tables store table identity and column names; table bindings expand with new rows. For named ranges store name and scope.

### Snapshot integrity

Fingerprint canonical types, dimensions, missing values, values and chart-relevant metadata:

\[
H = \operatorname{SHA256}(\text{schema}\parallel\text{values}\parallel\text{specification})
\]

## 8. Chart lifecycle and persistence

```text
ChartDefinition
  chart_id
  schema_version
  workbook_identity
  source_bindings
  chart_type
  transformations
  axes
  styles
  layout
  refresh_policy
  last_render_hash
  shape_identity
```

Use workbook custom XML parts for metadata where supported, plus portable JSON export. Do not rely solely on shape names, which users can change. Attach a chart identifier to managed shapes through documented metadata properties where practical.

### Transactional replacement

```text
Capture source snapshot
        ↓
Render new artifact
        ↓
Validate file and checksum
        ↓
Insert temporary new shape
        ↓
Apply position and dimensions
        ↓
Update registry reference
        ↓
Remove previous shape
        ↓
Commit success
```

Preserve the previous chart if a pre-commit step fails. Excel COM does not guarantee workbook-wide atomicity; use a recovery journal and reconciliation. Missing sources are marked unresolved, never silently rebound.

## 9. Native user interface

Initial Win32 chart designer sections:

- **Data:** Source ranges, X/Y series, categories, grouping.
- **Chart:** Family and subtype.
- **Axes:** Labels, bounds, scales, tick formatting.
- **Style:** Fonts, colors, markers, lines, legend.
- **Analysis:** Aggregation, trendline, confidence intervals when supported.
- **Output:** Dimensions, location, export format, refresh policy.

Validate inputs before dispatch. A later native COM Ribbon companion can expose Ribbon extensibility and invoke the same command service as the XLL.

## 10. Security model

Native code running in Excel requires strong controls.

### Trust boundaries

1. **Workbook → native host:** Treat values, formulas, names and metadata as untrusted.
2. **Native host → worker:** Validate every protocol field and message size.
3. **Worker → artifact manager:** Validate rendered files before Excel insertion.
4. **Installer/update → system:** Verify package integrity and signatures.

Run the worker with ordinary user privileges. Restrict named pipes to the current user and verify process identity as appropriate. Avoid unauthenticated localhost services. Never weaken Excel Trust Center settings automatically.

### Dependency policy

Pin versions, review licenses and vulnerabilities, and document purpose for every external dependency. Microsoft SDK, Windows SDK and CPython are platform dependencies; numerical libraries are optional and do not imply third-party cloud APIs.

## 11. Build and CI/CD system

### Build environment

- Visual Studio Build Tools / MSVC x64.
- Windows SDK.
- CMake and Ninja.
- Microsoft Excel XLL SDK.
- CPython and standard library.
- GitHub Actions Windows runners or equivalent CI.

### Native build targets

```text
xlfly_core
xlfly_excel_api
xlfly_ipc
xlfly_worker_supervisor
xlfly_artifact_manager
xlfly_xll
xlfly_native_tests
```

Pass Excel SDK paths through CMake configuration, not hard-coded machine paths. Build presets define architecture, compiler, SDK location, optimization and output directory.

### CI pipeline

1. Source validation — formatting, linting, dependencies.
2. Native build — CMake configure, compile and link.
3. Python tests — data, statistics, renderer.
4. Protocol tests — framing, schema, compatibility.
5. Security checks — static analysis, fuzzing, dependency audit.
6. Packaging — worker bundle, XLL, installer.
7. Integration tests — Excel-hosted tests on controlled Windows machines.

Compilation and standalone tests can run without Excel installed. Excel integration testing requires licensed Excel and an interactive desktop; do not assume hosted CI includes it.

## 12. Test strategy

- **Unit:** Geometry, statistics, formatting, serialization, native utilities.
- **Contract:** C++/Python agreement on framing, types, versions and errors.
- **Integration:** XLL load, registration, extraction, rendering, shape insertion.
- **Recovery:** Worker crashes, timeouts, workbook closure, protection, invalid SVG, interrupted replacement.
- **Compatibility:** Supported Microsoft 365 builds, update channels, Windows configurations, DPI and locale.

### Numerical validation

Compare calculations with independent reference values and use property-based tests.

For nonconstant valid `X`:

\[
\operatorname{corr}(X,X)=1
\]

For histogram bins covering all valid observations:

\[
\sum_i n_i=N
\]

Compare normalized SVG geometry and semantic properties, not just pixels.

### Performance

\[
T_{\text{total}}=T_{\text{read}}+T_{\text{serialize}}+T_{\text{IPC}}+T_{\text{compute}}+T_{\text{render}}+T_{\text{insert}}
\]

Measure peak memory and Excel responsiveness alongside latency. Set targets after a baseline prototype; do not promise million-point subsecond performance without evidence.

## 13. AAWDF — AI-agent implementation framework

Four governance roles:

1. **Registrar:** Create work items, assign IDs, record requirements and dependencies.
2. **Receiver:** Retrieve specifications, verify Microsoft documentation, define acceptance tests.
3. **Maker:** Implement source, documentation, tests and build artifacts.
4. **Checker:** Independently review, test and verify evidence.

### Specialist engineering agents

- **ARCH:** Architecture, interfaces, ADRs.
- **XLL:** Excel C API, registration, lifecycle, native memory.
- **COM:** Workbook automation, shapes, insertion.
- **IPC:** Named pipes, supervision, protocol.
- **PY:** Data, transforms, statistics.
- **SVG:** Geometry, layout, renderers, exports.
- **UI:** Native designer and Ribbon.
- **QA:** Unit, integration, recovery, performance.
- **SEC:** Security, dependencies, hardening.
- **REL:** Build, packaging, signing, releases.

Maker and Checker must be independent.

### Agent task specification

```yaml
task_id: XLFLY-M0-001
title: Implement native XLL lifecycle
owner: XLL
checker: QA
inputs:
  - docs/architecture/xll-lifecycle.md
  - docs/contracts/native-host.md
allowed_paths:
  - native/xll/**
  - tests/native/**
acceptance:
  - xlAutoOpen exported
  - xlAutoClose exported
  - x64 compilation succeeds
  - no exception escapes callbacks
  - lifecycle tests pass
evidence:
  - compiler_log
  - test_report
  - changed_files
  - source_revision
merge_policy:
  independent_checker: required
  human_approval: required
```

File ownership reduces conflicts between parallel agents.

### Execution ledger

```text
STARTED
INPUT_HASHED
TOOL_USE
DECISION
HUMAN_REVIEW
FINALIZED
```

Record role, task ID, revision, tool, outcome and artifact hashes; exclude secrets and sensitive workbook contents.

### Agent authority

Agents may implement, compile, test, refactor and open pull requests within assigned boundaries. Architecture changes, protocol major versions, new dependencies, releases and protected-branch changes require human approval.

Checker evidence states:

- **Verified:** Executed with recorded evidence.
- **Reviewed:** Examined without complete runtime verification.
- **Unverified:** Insufficient evidence.

## 14. Dependency-driven development plan

```text
M0 — Foundation and Contracts
    Repository · CMake · SDK · Architecture · Protocol
              ↓
    ┌───────────────────────────┐
    │ M1A — Native Host         │
    │ M1B — Python Engine       │
    └───────────────────────────┘
              ↓
M2 — End-to-End Integration
    Selected range → Python → Excel chart
              ↓
M3 — Persistence and Refresh
              ↓
M4 — Chart Catalogue and Designer
              ↓
M5 — Security, Compatibility and Release
```

### M0 — Foundation

Deliver reproducible x64 build, pinned SDK, source layout, ADRs, protocol schema and tests.

**Exit gate:** Clean build in documented Windows environment.

### M1 — Native and Python foundations

Native: lifecycle, `XLFLY.VERSION()`, registration, diagnostics, worker launcher. Python: typed data, worker protocol, scatter geometry, SVG.

**Exit gate:** Native compilation, deterministic SVG and IPC contract tests pass.

### M2 — First complete chart

Command reads two numeric columns and creates a scatter chart in Excel.

**Exit gate:** Real Excel insertion and persistence across save/reopen.

### M3 — Chart management

Chart IDs, metadata, source rebinding, manual refresh, replacement and recovery.

**Exit gate:** Edit and refresh without identity loss or workbook corruption.

### M4 — Chart catalogue

Line, column, bar, histogram, box plot, heatmap; chart designer and validation.

**Exit gate:** Numerical, geometry and integration tests pass for every type.

### M5 — Product release

Packaging, signing, installation, diagnostics, compatibility and recovery.

**Exit gate:** Reproducible release candidate validated on supported Excel.

## 15. Estimated engineering effort

Planning estimates, not measured durations:

- M0 — Foundation: **2–3 weeks**.
- M1 — Native and Python foundations: **3–5 weeks**.
- M2 — End-to-end chart: **3–4 weeks**.
- M3 — Persistence and refresh: **3–5 weeks**.
- M4 — Catalogue and designer: **5–8 weeks**.
- M5 — Hardening and release: **4–6 weeks**.

Sequential total: **approximately 20–31 engineering weeks**. Parallel agents may reduce elapsed time, but integration, review and Excel-hosted testing remain. Native Excel integration and workbook artifact lifecycle present the greatest schedule uncertainty.

## 16. Release engineering

```text
xlfly-release/
├── xlfly.xll
├── xlfly-engine.exe
├── runtime/
├── resources/
├── schemas/
├── licenses/
├── manifest.json
└── installer/
```

Installer requirements:

1. Detect Windows and Excel architecture.
2. Verify package integrity.
3. Install binaries in controlled location.
4. Register add-in through supported Excel mechanisms.
5. Preserve existing workbook content.
6. Support clean uninstall and repair.

Sign native binaries and installer. Do not silently change Trust Center policies.

### Versioning

```text
Product:       0.1.0
XLL:           0.1.0
Engine:        0.1.0
IPC:           1.0
Chart schema:  1.0
```

Negotiate XLL/worker protocol compatibility before jobs.

## 17. Principal failure modes

- **XLL initialization failure:** Safe load failure and diagnostics.
- **Worker crash:** Isolation, restart, job failure reporting.
- **Excel busy/modal:** Controlled retry or deferred insertion.
- **Workbook closed:** Cancel or discard stale result.
- **Source deleted:** Mark binding unresolved.
- **Protected worksheet:** Detect and report insertion restriction.
- **Invalid SVG:** Reject before insertion.
- **Protocol mismatch:** Refuse incompatible communication.
- **Memory leak:** Ownership tests and allocation instrumentation.
- **Corrupted registry:** Recover last valid metadata.
- **Multiple Excel instances:** Explicit instance/workbook identity.
- **Untrusted input:** Validation, size limits, no code execution.

Each failure mode requires automated testing where practical and a recovery procedure.

## 18. Definition of production readiness

- [ ] XLL loads/unloads reliably on supported Excel builds.
- [ ] Registered worksheet functions return correct values.
- [ ] Python worker starts, communicates and shuts down correctly.
- [ ] Chart creation handles valid and invalid data.
- [ ] Charts survive workbook save/reopen.
- [ ] Refresh avoids duplicate/orphaned shapes.
- [ ] Worker failure does not terminate Excel.
- [ ] Protocol and schema migrations tested.
- [ ] Installation/uninstallation repeatable.
- [ ] Release binaries signed and traceable to revisions.
- [ ] Independent Checker evidence recorded.
- [ ] User documentation and diagnostics complete.

## 19. Documentation package

```text
docs/
  00-product-requirements.md
  01-system-architecture.md
  02-excel-c-api-contracts.md
  03-xll-lifecycle.md
  04-memory-ownership.md
  05-threading-and-com.md
  06-ipc-protocol.md
  07-data-model.md
  08-python-engine.md
  09-svg-renderer.md
  10-chart-registry.md
  11-workbook-lifecycle.md
  12-security-model.md
  13-test-strategy.md
  14-build-and-release.md
  15-agent-governance.md
  16-operations-and-recovery.md
```

Official Microsoft references:

- [Developing Excel XLLs](https://learn.microsoft.com/en-us/office/client-developer/excel/developing-excel-xlls)
- [Creating XLLs](https://learn.microsoft.com/en-us/office/client-developer/excel/creating-xlls)
- [Data Types Used by Excel](https://learn.microsoft.com/en-us/office/client-developer/excel/data-types-used-by-excel)
- [Multithreaded Recalculation](https://learn.microsoft.com/en-us/office/client-developer/excel/multithreaded-recalculation-in-excel)
- [Asynchronous User-Defined Functions](https://learn.microsoft.com/en-us/office/client-developer/excel/asynchronous-user-defined-functions)
- [xlAutoOpen](https://learn.microsoft.com/en-us/office/client-developer/excel/xlautoopen)

## 20. Implementation directive for the agent team

### Initial engineering order — XLFLY-M0

Establish the xlfly repository as a native C++20 Excel XLL project with an independently executable Python engine.

Freeze architecture contracts, implement reproducible build system, register `XLFLY.VERSION()`, establish named-pipe protocol and create automated contract tests.

Every task must have a Maker, independent Checker, acceptance criteria and reproducible evidence.

**Do not begin workbook chart insertion until native lifecycle and worker communication gates have passed.**

### Final architectural position

xlfly is an Excel-native extension platform with a proprietary Python analytical and visualization engine. The XLL provides host integration; the worker provides isolated computation; the workbook artifact manager provides persistent chart lifecycle management.

The first critical transaction is:

\[
\boxed{\text{Excel Range}\rightarrow\text{Native XLL}\rightarrow\text{Python Engine}\rightarrow\text{SVG}\rightarrow\text{Excel Shape}}
\]

Once this transaction is reliable, extensible and recoverable, additional chart families become independent engineering modules rather than changes to the core architecture.