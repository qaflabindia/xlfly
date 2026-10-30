# Complete xlfly source walkthrough

Companion to Building a local Python add in for Excel. Source snapshot 27b4fcaf3c92b6b7fa45e5db57631505bdc4f330. All tracked text files under addin are reproduced below without abbreviation. Paths are relative to addin. Binary icons are preserved in the source archive. No node_modules, generated dist, vendor packages or machine files are included.

Read in this order: configuration → value and RPC contracts → host client and worker → Python runtime and API → Excel services and trust → application and UI → build and tests. Each explanation tells you what to watch for before the complete listing.

## Sequential construction prompts

Use these in order in one project. Each stage consumes the previous files and the complete reference listings below. Ask the agent to return complete changed files, an explicit contract summary, test results and the next-stage handoff. These are instructional prompts authored for this walkthrough.

### Stage 1 Create the shell

> Create package.json, tsconfig.json, manifest.xml, src/functions/functions.json and a minimal taskpane.html. Use the snapshot dependency versions and lockfile when reproducing xlfly. Give a new product a new GUID and branding. Wire VERSION first; preserve PY and CALL metadata for later stages. Include a gitignore. Verify the build and distinguish loading a page from executing Python.

### Stage 2 Define values and messages

> Implement src/shared/values.ts, python/xlfly/_types.py and src/shared/rpc.ts using the complete reference listings below. Keep the same protocol v1 and wire format. Implement requests, replies, notifications, pending-call cleanup, version refusal and timeouts. Test blanks, errors, scalar unwrapping, ragged matrices and bidirectional requests before integrating a worker.

### Stage 3 Build the Python package

> Implement _context.py, _host.py, _registry.py, _runtime.py, _api.py and __init__.py. Preserve JSON entry points, the registry catalog shape, function/command separation and restoration of context variables. Compile the final expression separately, reject awaitable formula results and support async commands through an injected bridge. Test all behavior with a fake host; do not depend on Excel yet.

### Stage 4 Connect the worker

> Implement worker.ts, kernel-core.ts, kernel-client.ts and batcher.ts against the existing contracts. Load Pyodide in a module worker, install the Python bundle in its virtual filesystem and expose the documented RPC methods. Retain module sources for recovery, validate live command tokens at the host and serialize commands and events. Add the real-Pyodide integration test with fake Excel services. Verify 1+1, scalar and matrix calls, missing names, stale tokens and timeout recovery.

### Stage 5 Add workbook storage and services

> Implement codestore-format.ts, office-store.ts and excel-services.ts. Preserve the versioned custom XML representation, length-prefixed trust hashing and separate trust storage. Add bulk reads and writes, scalar Excel function calls, native charts and PNG images behind the host guard. Preserve cell limits. Explain the self-write tracker heuristic and leave real Office.js validation explicit.

### Stage 6 Join the application and interface

> Implement app.ts and taskpane.ts, HTML and CSS using the wireframe and the public methods established by previous stages. Associate VERSION, PY and CALL once. Load stored code, refresh trust, gate formula evaluation and register events. Provide review/trust, draft module editing, save/reload, catalogs, restart and logs. Preserve dirty drafts during status updates. Validate states without claiming mocked UI tests prove Excel compatibility.

### Stage 7 Package and verify the whole system

> Implement build.mjs and scripts for package fetching, HTTPS serving and Mac sideloading using the reference files. Preserve the correct host IIFE and worker ESM outputs, runtime asset paths and build-time URL substitution. Keep Windows and web sideload instructions separate. Run typecheck, unit tests, Python tests, real-kernel integration and build. Deliver a full file inventory, exact commands, actual results and remaining real-host checks.

## Complete file listings

## 1 package.json

Defines the build and test entry points and direct npm dependencies. Read the scripts before running commands: check chains five stages; start builds and serves; serve alone does not rebuild. The private and license fields are preserved from the project.

Original size: 1324 bytes. SHA-256: `1e9fa2252d13c1664ee48fecba16bf54db4f1debda42956317dcd7e1ec046b04`.

````json
{
  "name": "xlfly-addin",
  "version": "0.1.0",
  "description": "xlfly: local Python for Excel (Office Add-in for Mac, Windows and web)",
  "scripts": {
    "build": "node build.mjs",
    "typecheck": "tsc -p tsconfig.json",
    "test": "node --test tests/*.test.ts",
    "test:py": "cd python && python3 -m unittest discover -s tests -t . -v",
    "test:kernel": "node --test tests/kernel.integration.ts",
    "check": "npm run typecheck && npm test && npm run test:py && npm run test:kernel && npm run build",
    "fetch:pyodide": "node scripts/fetch-pyodide-packages.mjs",
    "certs": "office-addin-dev-certs install",
    "start": "npm run build && node scripts/serve.mjs",
    "sideload:mac": "bash scripts/sideload-mac.sh",
    "validate": "office-addin-manifest validate dist/manifest.xml",
    "setup:mac": "bash scripts/setup-mac.sh",
    "serve": "node scripts/serve.mjs"
  },
  "license": "UNLICENSED",
  "devDependencies": {
    "@types/custom-functions-runtime": "^1.6.13",
    "@types/node": "^26.6.6",
    "@types/office-js": "^1.0.615",
    "@types/office-runtime": "^1.0.40",
    "esbuild": "^0.28.2",
    "office-addin-dev-certs": "^3.0.1",
    "office-addin-manifest": "^3.0.1",
    "typescript": "^7.0.2"
  },
  "dependencies": {
    "pyodide": "^314.0.7"
  },
  "private": true,
  "type": "module"
}
````

## 2 tsconfig.json

Checks the source and tests without emitting JavaScript. Bundler resolution and allowImportingTsExtensions support the explicit .ts imports. DOM and WebWorker libraries describe both execution environments; esbuild handles output generation.

Original size: 441 bytes. SHA-256: `77fe516077bf3377a28c9624459902f18f1a9a5c7084a59d81876d6d07ed01e3`.

````json
{
  "compilerOptions": {
    "target": "ES2020",
    "module": "ESNext",
    "moduleResolution": "Bundler",
    "lib": ["ES2022", "DOM", "WebWorker"],
    "strict": true,
    "noEmit": true,
    "allowImportingTsExtensions": true,
    "verbatimModuleSyntax": true,
    "erasableSyntaxOnly": true,
    "skipLibCheck": true,
    "types": ["office-js", "office-runtime", "custom-functions-runtime", "node"]
  },
  "include": ["src", "tests"]
}
````

## 3 manifest.xml

Declares the add-in identity, workbook permissions and SharedRuntime 1.1 requirement. Follow Taskpane.Url through Runtimes, the custom-function Page and the ribbon action: they point at one page. The build replaces BASE_URL placeholders; XML resource IDs must match their references.

Original size: 4933 bytes. SHA-256: `cba5828efbace29c392462d86720138dc3fb900905faeae1b46ac2d2324abe4e`.

````xml
<?xml version="1.0" encoding="UTF-8"?>
<!-- xlfly Office Add-in manifest. {{BASE_URL}} is replaced at build time (XLFLY_BASE_URL, default https://localhost:3000). -->
<OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
           xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
           xmlns:bt="http://schemas.microsoft.com/office/officeappbasictypes/1.0"
           xsi:type="TaskPaneApp">
  <Id>6f1c2b8e-4d3a-4f5e-9b7c-2a1d0e9f8c71</Id>
  <Version>0.1.0.0</Version>
  <ProviderName>QAF Lab</ProviderName>
  <DefaultLocale>en-US</DefaultLocale>
  <DisplayName DefaultValue="xlfly"/>
  <Description DefaultValue="Local Python for Excel: worksheet functions, commands and charts, running on your computer."/>
  <IconUrl DefaultValue="{{BASE_URL}}/assets/icon-32.png"/>
  <HighResolutionIconUrl DefaultValue="{{BASE_URL}}/assets/icon-64.png"/>
  <SupportUrl DefaultValue="https://github.com/qaflabindia/xlfly"/>
  <AppDomains>
    <AppDomain>{{BASE_URL}}</AppDomain>
  </AppDomains>
  <Hosts>
    <Host Name="Workbook"/>
  </Hosts>
  <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="SharedRuntime" MinVersion="1.1"/>
    </Sets>
  </Requirements>
  <DefaultSettings>
    <SourceLocation DefaultValue="{{BASE_URL}}/taskpane.html"/>
  </DefaultSettings>
  <Permissions>ReadWriteDocument</Permissions>
  <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
    <Hosts>
      <Host xsi:type="Workbook">
        <Runtimes>
          <Runtime resid="Taskpane.Url" lifetime="long"/>
        </Runtimes>
        <AllFormFactors>
          <ExtensionPoint xsi:type="CustomFunctions">
            <Script>
              <SourceLocation resid="Functions.Script.Url"/>
            </Script>
            <Page>
              <SourceLocation resid="Taskpane.Url"/>
            </Page>
            <Metadata>
              <SourceLocation resid="Functions.Metadata.Url"/>
            </Metadata>
            <Namespace resid="Functions.Namespace"/>
          </ExtensionPoint>
        </AllFormFactors>
        <DesktopFormFactor>
          <GetStarted>
            <Title resid="GetStarted.Title"/>
            <Description resid="GetStarted.Description"/>
            <LearnMoreUrl resid="LearnMore.Url"/>
          </GetStarted>
          <FunctionFile resid="Taskpane.Url"/>
          <ExtensionPoint xsi:type="PrimaryCommandSurface">
            <OfficeTab id="TabHome">
              <Group id="Xlfly.Group">
                <Label resid="Group.Label"/>
                <Icon>
                  <bt:Image size="16" resid="Icon.16"/>
                  <bt:Image size="32" resid="Icon.32"/>
                  <bt:Image size="80" resid="Icon.80"/>
                </Icon>
                <Control xsi:type="Button" id="Xlfly.TaskpaneButton">
                  <Label resid="Button.Label"/>
                  <Supertip>
                    <Title resid="Button.Label"/>
                    <Description resid="Button.Tooltip"/>
                  </Supertip>
                  <Icon>
                    <bt:Image size="16" resid="Icon.16"/>
                    <bt:Image size="32" resid="Icon.32"/>
                    <bt:Image size="80" resid="Icon.80"/>
                  </Icon>
                  <Action xsi:type="ShowTaskpane">
                    <TaskpaneId>Xlfly.Taskpane</TaskpaneId>
                    <SourceLocation resid="Taskpane.Url"/>
                  </Action>
                </Control>
              </Group>
            </OfficeTab>
          </ExtensionPoint>
        </DesktopFormFactor>
      </Host>
    </Hosts>
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16" DefaultValue="{{BASE_URL}}/assets/icon-16.png"/>
        <bt:Image id="Icon.32" DefaultValue="{{BASE_URL}}/assets/icon-32.png"/>
        <bt:Image id="Icon.80" DefaultValue="{{BASE_URL}}/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Functions.Script.Url" DefaultValue="{{BASE_URL}}/app.js"/>
        <bt:Url id="Functions.Metadata.Url" DefaultValue="{{BASE_URL}}/functions.json"/>
        <bt:Url id="Taskpane.Url" DefaultValue="{{BASE_URL}}/taskpane.html"/>
        <bt:Url id="LearnMore.Url" DefaultValue="https://github.com/qaflabindia/xlfly"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="Functions.Namespace" DefaultValue="XLFLY"/>
        <bt:String id="GetStarted.Title" DefaultValue="xlfly is ready"/>
        <bt:String id="Group.Label" DefaultValue="xlfly"/>
        <bt:String id="Button.Label" DefaultValue="Python"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="GetStarted.Description" DefaultValue="Open the xlfly pane from the Home tab to write Python for this workbook."/>
        <bt:String id="Button.Tooltip" DefaultValue="Open the xlfly Python pane."/>
      </bt:LongStrings>
    </Resources>
  </VersionOverrides>
</OfficeApp>
````

## 4 src/functions/functions.json

Declares the stable Excel entry points VERSION, PY and CALL. PY accepts code; CALL accepts a registered Python name. Repeating optional matrix arguments support values and ranges. The result matrix enables spill behavior; Python registration does not rewrite this metadata.

Original size: 1418 bytes. SHA-256: `63690e2dc7aa5bf50cf3f617e25145723c920dae5a734e6a56c90c9c2f5b079b`.

````json
{
  "functions": [
    {
      "id": "VERSION",
      "name": "VERSION",
      "description": "Returns the xlfly add-in version.",
      "result": { "type": "string", "dimensionality": "scalar" },
      "parameters": []
    },
    {
      "id": "PY",
      "name": "PY",
      "description": "Evaluates Python code. Arguments are available as arg1, arg2, ... and args. The value of the last expression is returned.",
      "helpUrl": "https://github.com/qaflabindia/xlfly",
      "result": { "type": "any", "dimensionality": "matrix" },
      "parameters": [
        { "name": "code", "description": "Python code.", "type": "string", "dimensionality": "scalar" },
        { "name": "args", "description": "Values or ranges passed to the code.", "type": "any", "dimensionality": "matrix", "optional": true, "repeating": true }
      ]
    },
    {
      "id": "CALL",
      "name": "CALL",
      "description": "Calls a Python function registered with @xlfly.func in this workbook.",
      "helpUrl": "https://github.com/qaflabindia/xlfly",
      "result": { "type": "any", "dimensionality": "matrix" },
      "parameters": [
        { "name": "name", "description": "Function name.", "type": "string", "dimensionality": "scalar" },
        { "name": "args", "description": "Values or ranges passed to the function.", "type": "any", "dimensionality": "matrix", "optional": true, "repeating": true }
      ]
    }
  ]
}
````

## 5 src/shared/values.ts

Defines the TypeScript half of the wire contract. argToWire handles custom-function inputs; rangeToWire uses valueTypes to distinguish error cells from text. wireToMatrix preserves errors for formula mapping, whereas wireToRangeValues writes them as text. Both pad rows. parseAddress resolves quoted sheet names and unqualified ranges.

Original size: 3652 bytes. SHA-256: `3a9f9b2172dea928cf08a6af26c9f00e70c1d021a8b037dd032b9a5e278655a5`.

````typescript
// Conversion between Office.js values and the xlfly wire format (see python/xlfly/_types.py).

export type WireScalar = number | string | boolean | null | { $err: string };
export type Wire = WireScalar | WireScalar[][];

export const ERROR_CODES = new Set([
  "#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A",
  "#GETTING_DATA", "#SPILL!", "#CALC!", "#BUSY!",
]);

export function isWireError(value: unknown): value is { $err: string } {
  return typeof value === "object" && value !== null && !Array.isArray(value) && typeof (value as any).$err === "string";
}

/** A custom-function argument (scalar or matrix) to wire format. */
export function argToWire(value: unknown): Wire {
  if (Array.isArray(value)) {
    return value.map((row) => (Array.isArray(row) ? row : [row]).map(scalarToWire));
  }
  return scalarToWire(value);
}

function scalarToWire(value: unknown): WireScalar {
  if (value === undefined || value === null) return null;
  if (typeof value === "number" || typeof value === "boolean") return value;
  if (typeof value === "string") return value === "" ? null : value;
  // CustomFunctions.Error passed as input when errors are allowed through
  if (typeof value === "object" && typeof (value as any).code === "string" && ERROR_CODES.has((value as any).code)) {
    return { $err: (value as any).code };
  }
  return String(value);
}

/** Range.values + Range.valueTypes to wire format. */
export function rangeToWire(values: unknown[][], valueTypes?: string[][]): WireScalar[][] {
  return values.map((row, r) =>
    row.map((cell, c) => {
      const type = valueTypes?.[r]?.[c];
      if (type === "Error" && typeof cell === "string") return { $err: cell };
      if (type === "Empty" || cell === "") return null;
      return scalarToWire(cell);
    }),
  );
}

/** Wire 2-D array to values accepted by Range.values. Errors are written as their text form. */
export function wireToRangeValues(wire: Wire): (string | number | boolean)[][] {
  const rows = Array.isArray(wire) ? wire : [[wire]];
  const width = Math.max(1, ...rows.map((r) => r.length));
  return rows.map((row) => {
    const out = row.map((cell): string | number | boolean => {
      if (cell === null) return "";
      if (isWireError(cell)) return cell.$err;
      return cell;
    });
    while (out.length < width) out.push("");
    return out;
  });
}

export interface ResultError {
  error: { code: string; message: string };
}

/** Wire result to a custom-function matrix result, with errors returned separately for mapping. */
export function wireToMatrix(wire: Wire): (string | number | boolean | ResultError)[][] {
  const rows = Array.isArray(wire) ? wire : [[wire]];
  const width = Math.max(1, ...rows.map((r) => r.length));
  return rows.map((row) => {
    const out = row.map((cell): string | number | boolean | ResultError => {
      if (cell === null) return "";
      if (isWireError(cell)) return { error: { code: cell.$err, message: "" } };
      return cell;
    });
    while (out.length < width) out.push("");
    return out;
  });
}

/** Parse "Sheet1!A1:B2", "'My Sheet'!A1" or "A1" into sheet and range parts. */
export function parseAddress(address: string): { sheet: string | null; range: string } {
  const trimmed = address.trim();
  const bang = trimmed.lastIndexOf("!");
  if (bang < 0) return { sheet: null, range: trimmed };
  let sheet = trimmed.slice(0, bang);
  if (sheet.startsWith("'") && sheet.endsWith("'")) sheet = sheet.slice(1, -1).replace(/''/g, "'");
  const range = trimmed.slice(bang + 1);
  if (!sheet || !range) throw new Error(`invalid address: ${address}`);
  return { sheet, range };
}
````

## 6 python/xlfly/_types.py

Defines the matching Python conversion layer. from_wire creates XLError objects; to_wire turns lists, arrays and DataFrames into serializable cells. One-dimensional output becomes a column, ragged rows are padded and non-finite numbers become #NUM!. unwrap_single explains scalar versus range inputs. Date helpers use one epoch and strip timezone metadata.

Original size: 5137 bytes. SHA-256: `498852396965474392dcf3c490569b1dbe5fcf89d2e07032c4697b7546dda7a5`.

````python
"""Excel <-> Python value conversion.

Wire format (JSON, shared with the TypeScript host):
  number | string | bool | null       scalar cell value
  {"$err": "#N/A"}                    Excel error
  [[...], [...]]                      2-D array, row-major
"""

from __future__ import annotations

import datetime as _dt
import math
from typing import Any

ERROR_CODES = frozenset(
    {"#NULL!", "#DIV/0!", "#VALUE!", "#REF!", "#NAME?", "#NUM!", "#N/A", "#GETTING_DATA", "#SPILL!", "#CALC!", "#BUSY!"}
)

_EXCEL_EPOCH = _dt.datetime(1899, 12, 30)


class XLError(Exception):
    """An Excel error value such as #N/A. Can be returned or raised."""

    def __init__(self, code: str = "#VALUE!", message: str = ""):
        if code not in ERROR_CODES:
            raise ValueError(f"unknown Excel error code: {code!r}")
        super().__init__(message or code)
        self.code = code
        self.message = message

    def __eq__(self, other: object) -> bool:
        return isinstance(other, XLError) and other.code == self.code

    def __hash__(self) -> int:
        return hash(self.code)

    def __repr__(self) -> str:
        return f"XLError({self.code!r})"


def to_excel_date(value: _dt.date | _dt.datetime) -> float:
    """Convert a date or datetime to an Excel serial number (1900 date system)."""
    if not isinstance(value, _dt.datetime):
        value = _dt.datetime(value.year, value.month, value.day)
    if value.tzinfo is not None:
        value = value.replace(tzinfo=None)
    delta = value - _EXCEL_EPOCH
    return delta.days + delta.seconds / 86400 + delta.microseconds / 86_400_000_000


def from_excel_date(serial: float) -> _dt.datetime:
    """Convert an Excel serial number (1900 date system) to a datetime."""
    return _EXCEL_EPOCH + _dt.timedelta(days=float(serial))


# ---------------------------------------------------------------- wire -> Python

def from_wire(value: Any) -> Any:
    if isinstance(value, dict):
        if "$err" in value:
            return XLError(value["$err"])
        raise ValueError(f"unexpected object on the wire: {value!r}")
    if isinstance(value, list):
        return [from_wire(v) for v in value]
    return value


# ---------------------------------------------------------------- Python -> wire

def _scalar_to_wire(value: Any) -> Any:
    if value is None:
        return None
    if isinstance(value, XLError):
        return {"$err": value.code}
    if isinstance(value, bool):
        return value
    if isinstance(value, (int, float)):
        f = float(value)
        if math.isnan(f) or math.isinf(f):
            return {"$err": "#NUM!"}
        return int(value) if isinstance(value, int) and abs(value) < 2**53 else f
    if isinstance(value, str):
        return value
    if isinstance(value, (_dt.datetime, _dt.date)):
        return to_excel_date(value)
    # numpy / pandas scalars without importing them
    item = getattr(value, "item", None)
    if callable(item) and getattr(value, "shape", None) == ():
        return _scalar_to_wire(item())
    type_name = type(value).__name__
    if type_name in ("NaTType", "NAType"):
        return None
    if type_name == "Timestamp":
        return to_excel_date(value.to_pydatetime())
    return str(value)


def _rectangular(rows: list[list[Any]]) -> list[list[Any]]:
    width = max((len(r) for r in rows), default=0)
    if width == 0:
        return [[None]]
    return [r + [None] * (width - len(r)) for r in rows]


def to_wire(value: Any, header: bool = True) -> Any:
    """Convert a Python result to the wire format.

    Scalars stay scalars. 1-D sequences become a single column.
    DataFrames include a header row when ``header`` is true.
    """
    type_name = type(value).__name__
    if type_name == "DataFrame":
        rows = value.values.tolist()
        body = [[_scalar_to_wire(v) for v in r] for r in rows]
        if header:
            body.insert(0, [_scalar_to_wire(c) for c in value.columns])
        return _rectangular(body) if body else [[None]]
    if type_name == "Series":
        return [[_scalar_to_wire(v)] for v in value.tolist()] or [[None]]
    if type_name == "ndarray":
        if value.ndim == 0:
            return _scalar_to_wire(value.item())
        if value.ndim == 1:
            return [[_scalar_to_wire(v)] for v in value.tolist()] or [[None]]
        if value.ndim == 2:
            return _rectangular([[_scalar_to_wire(v) for v in r] for r in value.tolist()])
        raise ValueError("arrays with more than 2 dimensions cannot be returned to Excel")
    if isinstance(value, (list, tuple)):
        if not value:
            return [[None]]
        if all(isinstance(r, (list, tuple)) for r in value):
            return _rectangular([[_scalar_to_wire(v) for v in r] for r in value])
        return [[_scalar_to_wire(v)] for v in value]
    return _scalar_to_wire(value)


def unwrap_single(value: Any) -> Any:
    """Return the scalar inside a 1x1 array, otherwise the value unchanged."""
    if isinstance(value, list) and len(value) == 1 and isinstance(value[0], list) and len(value[0]) == 1:
        return value[0][0]
    return value
````

## 7 src/shared/rpc.ts

Implements a symmetric request channel. request records a pending promise before sending; handle routes responses back by ID and incoming requests by method. Timeouts remove pending entries, so late responses are ignored. close rejects every pending operation. The deadline is metadata, not automatic execution cancellation.

Original size: 5003 bytes. SHA-256: `9e2aa73db05c2e0a0948daa8a78e2961a1139fedc8571f1db400f0ce9eb7ca2b`.

````typescript
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
````

## 8 src/host/batcher.ts

Assigns an ID to each worksheet call and flushes the queue on a timer or at its size threshold. Results are matched by ID rather than order. A missing result rejects only that call; a transport failure rejects the batch. Batching reduces message traffic without making Python evaluation parallel.

Original size: 1846 bytes. SHA-256: `b96dd28a86744d08cd85ff971de12cb7c2b0f88020e5bf23c40ff062fbe760ed`.

````typescript
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
````

## 9 src/host/kernel-client.ts

Owns worker startup, RPC service registration and recovery. ensureStarted shares a startup promise; loadModules retains source for a later restart. guard validates live command tokens. withToken serializes commands and events and removes tokens in finally. Function or command timeout stops the worker; the next request starts it again.

Original size: 6922 bytes. SHA-256: `bb26c0b5069e6730ba2d19b10e8e60181f1a54578d2c8db89cbd2142d62becf8`.

````typescript
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
````

## 10 src/kernel/worker.ts

The thin worker entry point installs an RPC endpoint before startup. Its start handler dynamically imports Pyodide from the configured runtime URL and fetches the Python source bundle. Keeping this wrapper small lets kernel-core run under Node in integration tests too.

Original size: 923 bytes. SHA-256: `f20b31025b680bb146143f6b29be01b0f4f16c5b14b232158fbf33ce8cc326a5`.

````typescript
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
````

## 11 src/kernel/kernel-core.ts

Loads Pyodide, forwards stdout and stderr notifications, writes the Python package into the virtual filesystem and installs the Python-to-host bridge. It exposes load_modules, call_functions, run_command and event over RPC. JSON strings carry results across the Python boundary; package-import loading is attempted before evaluation.

Original size: 3010 bytes. SHA-256: `061b2ee55f40b98dd60cd83b853b73773250f453dd69aa085d2fbc36926bc4bf`.

````typescript
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
````

## 12 python/xlfly/_context.py

Uses context variables to track function or command execution and the current token. The running context manager restores previous values even after an error. require_command provides the Python-side rule; it does not replace the independent host token check.

Original size: 1501 bytes. SHA-256: `58e0785cb736fb44a589b9a54620b11c923fc3629ee2c6f690a6bb4dc3eb5a40`.

````python
"""Execution contexts.

function  -- running a worksheet function; Excel is waiting on the result.
             Calls back into Excel are refused (they would deadlock and
             Excel forbids side effects from worksheet functions).
command   -- running a command or event handler; workbook access allowed.
"""

from __future__ import annotations

import contextvars
from contextlib import contextmanager
from typing import Iterator

FUNCTION = "function"
COMMAND = "command"

current_context: contextvars.ContextVar[str | None] = contextvars.ContextVar("xlfly_context", default=None)
current_token: contextvars.ContextVar[str | None] = contextvars.ContextVar("xlfly_token", default=None)


class ContextError(RuntimeError):
    """Raised when Excel is accessed from a context that does not allow it."""


@contextmanager
def running(kind: str, token: str | None = None) -> Iterator[None]:
    c = current_context.set(kind)
    t = current_token.set(token)
    try:
        yield
    finally:
        current_context.reset(c)
        current_token.reset(t)


def require_command(operation: str) -> str:
    kind = current_context.get()
    if kind != COMMAND:
        where = "a worksheet function" if kind == FUNCTION else "outside a command"
        raise ContextError(
            f"{operation} is not allowed from {where}; use @xlfly.command or an event handler"
        )
    token = current_token.get()
    if not token:
        raise ContextError("missing command token")
    return token
````

## 13 python/xlfly/_host.py

Adapts Python API calls into bridge requests. It checks command context, JSON-encodes parameters, awaits the installed bridge and decodes results. JavaScript failures are translated into HostError. A fake bridge makes this layer testable without Excel.

Original size: 1262 bytes. SHA-256: `bf26c6a63848dcac73c1d3c20cdfbea3bb16979c23c9a1f7ec411fd75fdb043f`.

````python
"""Bridge from Python to the Excel host.

The kernel installs a bridge: ``async bridge(method, params_json, token) -> result_json``.
In Pyodide it wraps a JavaScript function returning a Promise; in tests it is
a plain coroutine function.
"""

from __future__ import annotations

import json
from typing import Any, Awaitable, Callable

from ._context import require_command

Bridge = Callable[[str, str, str], Awaitable[str]]

_bridge: Bridge | None = None


def install_bridge(bridge: Bridge | None) -> None:
    global _bridge
    _bridge = bridge


class HostError(RuntimeError):
    def __init__(self, code: str, message: str):
        super().__init__(f"{code}: {message}")
        self.code = code


async def request(method: str, params: dict[str, Any]) -> Any:
    token = require_command(method)
    if _bridge is None:
        raise HostError("NO_HOST", "xlfly is not running inside Excel")
    try:
        raw = await _bridge(method, json.dumps(params), token)
    except Exception as exc:  # JavaScript errors surface as JsException in Pyodide
        message = str(exc)
        code = message.split(":", 1)[0] if ":" in message else "HOST_ERROR"
        raise HostError(code.strip(), message) from None
    return json.loads(raw) if raw else None
````

## 14 python/xlfly/_registry.py

Decorators register functions and commands under normalized uppercase names; on accepts only the known event types. clear resets catalogs during reload. These dictionaries are Python-side catalogs, not dynamically generated Excel metadata.

Original size: 1979 bytes. SHA-256: `1e19db5bd45f425b96fa5465d3bfff82e4a89b06562ce3b857ccf1ab1e256d47`.

````python
"""Registries filled by the decorators in user modules."""

from __future__ import annotations

from dataclasses import dataclass, field
from typing import Any, Callable

EVENT_TYPES = frozenset({"sheet_change", "selection_change", "workbook_open"})


@dataclass
class Registry:
    functions: dict[str, Callable[..., Any]] = field(default_factory=dict)
    commands: dict[str, Callable[..., Any]] = field(default_factory=dict)
    events: dict[str, list[Callable[..., Any]]] = field(default_factory=dict)

    def clear(self) -> None:
        self.functions.clear()
        self.commands.clear()
        self.events.clear()


REGISTRY = Registry()


def _key(name: str) -> str:
    key = name.strip().upper()
    if not key or not all(ch.isalnum() or ch in "._" for ch in key):
        raise ValueError(f"invalid name: {name!r}")
    return key


def _decorator(target: dict[str, Callable[..., Any]], name_or_fn: Any):
    def register(fn: Callable[..., Any], name: str | None = None):
        target[_key(name or fn.__name__)] = fn
        return fn

    if callable(name_or_fn):
        return register(name_or_fn)
    return lambda fn: register(fn, name_or_fn)


def func(name_or_fn: Any = None):
    """Expose a function to worksheets: ``=XLFLY.CALL("NAME", args...)``.

    Worksheet functions must be synchronous and side-effect free.
    """
    return _decorator(REGISTRY.functions, name_or_fn)


def command(name_or_fn: Any = None):
    """Expose a command, run from the task pane. May be ``async`` and may use the workbook API."""
    return _decorator(REGISTRY.commands, name_or_fn)


def on(event: str):
    """Subscribe a handler to a workbook event: sheet_change, selection_change, workbook_open."""
    if event not in EVENT_TYPES:
        raise ValueError(f"unknown event {event!r}; expected one of {sorted(EVENT_TYPES)}")

    def register(fn: Callable[..., Any]):
        REGISTRY.events.setdefault(event, []).append(fn)
        return fn

    return register
````

## 15 python/xlfly/_runtime.py

Coordinates module loading, compilation and execution. load_modules resets the registry and user package and reports errors per module; execution is not a transactional rollback of partially registered code. _compile_cell separates the trailing expression and keeps a 256-entry cache. call_functions sets FUNCTION context; commands and events use COMMAND context and await async results.

Original size: 6886 bytes. SHA-256: `8bdb5d8f7069fa51e17a50e8aca7a29e40bb0bc3662e427f391ecc8577ea0245`.

````python
"""Kernel entry points called by the JavaScript worker.

Every entry point takes and returns JSON strings so that no Python or
JavaScript proxies cross the boundary.
"""

from __future__ import annotations

import ast
import inspect
import json
import sys
import traceback
import types
from collections import OrderedDict
from typing import Any

import xlfly

from ._context import COMMAND, FUNCTION, running
from ._registry import REGISTRY
from ._types import XLError, from_wire, to_wire, unwrap_single

USER_PACKAGE = "xlfly_user"
_CODE_CACHE: "OrderedDict[str, tuple[Any, Any]]" = OrderedDict()
_CODE_CACHE_SIZE = 256


def _error(exc: BaseException, include_trace: bool = False) -> dict[str, Any]:
    if isinstance(exc, XLError):
        return {"code": exc.code, "message": exc.message or exc.code}
    message = f"{type(exc).__name__}: {exc}"
    error: dict[str, Any] = {"code": "#VALUE!", "message": message}
    if include_trace:
        error["trace"] = "".join(traceback.format_exception(type(exc), exc, exc.__traceback__, limit=-8))
    return error


def describe() -> dict[str, Any]:
    def info(fn: Any) -> dict[str, Any]:
        doc = inspect.getdoc(fn) or ""
        try:
            signature = str(inspect.signature(fn))
        except (TypeError, ValueError):
            signature = "(...)"
        return {"signature": signature, "doc": doc.splitlines()[0] if doc else ""}

    return {
        "functions": {name: info(fn) for name, fn in sorted(REGISTRY.functions.items())},
        "commands": {name: info(fn) for name, fn in sorted(REGISTRY.commands.items())},
        "events": sorted(REGISTRY.events),
    }


def load_modules(modules_json: str) -> str:
    """Replace all user modules. Input: [{"name": str, "source": str}, ...]."""
    modules = json.loads(modules_json)
    REGISTRY.clear()
    _CODE_CACHE.clear()
    for name in [n for n in sys.modules if n == USER_PACKAGE or n.startswith(USER_PACKAGE + ".")]:
        del sys.modules[name]
    package = types.ModuleType(USER_PACKAGE)
    package.__path__ = []
    sys.modules[USER_PACKAGE] = package

    errors = []
    for module in modules:
        name = module["name"]
        if not name.isidentifier():
            errors.append({"module": name, "code": "#NAME?", "message": "module name must be a Python identifier"})
            continue
        qualified = f"{USER_PACKAGE}.{name}"
        mod = types.ModuleType(qualified)
        mod.__file__ = f"<xlfly:{name}>"
        sys.modules[qualified] = mod
        setattr(package, name, mod)
        try:
            exec(compile(module["source"], mod.__file__, "exec"), mod.__dict__)
        except BaseException as exc:  # a user module must never take the kernel down
            if isinstance(exc, (KeyboardInterrupt, SystemExit)):
                exc = RuntimeError(f"module raised {type(exc).__name__}")
            errors.append({"module": name, **_error(exc, include_trace=True)})
    return json.dumps({**describe(), "errors": errors})


def _compile_cell(code: str) -> tuple[Any, Any]:
    cached = _CODE_CACHE.get(code)
    if cached is not None:
        _CODE_CACHE.move_to_end(code)
        return cached
    tree = ast.parse(code, filename="<cell>", mode="exec")
    last = None
    if tree.body and isinstance(tree.body[-1], ast.Expr):
        last = compile(ast.Expression(tree.body.pop().value), "<cell>", "eval")
    body = compile(tree, "<cell>", "exec")
    _CODE_CACHE[code] = (body, last)
    if len(_CODE_CACHE) > _CODE_CACHE_SIZE:
        _CODE_CACHE.popitem(last=False)
    return body, last


def evaluate(code: str, args: list[Any]) -> Any:
    """Run cell code; the value of a trailing expression is the result."""
    body, last = _compile_cell(code)
    namespace: dict[str, Any] = {
        "__name__": "__xlfly_cell__",
        "xlfly": xlfly,
        "user": sys.modules.get(USER_PACKAGE),
        "args": args,
    }
    for index, value in enumerate(args, start=1):
        namespace[f"arg{index}"] = value
    exec(body, namespace)
    return eval(last, namespace) if last is not None else namespace.get("result")


def _invoke(item: dict[str, Any]) -> Any:
    args = [unwrap_single(from_wire(a)) for a in item.get("args", [])]
    if item["kind"] == "eval":
        value = evaluate(item["code"], args)
    elif item["kind"] == "call":
        fn = REGISTRY.functions.get(str(item["name"]).upper())
        if fn is None:
            raise XLError("#NAME?", f"no xlfly function named {item['name']!r}")
        value = fn(*args)
    else:
        raise ValueError(f"unknown call kind {item['kind']!r}")
    if inspect.isawaitable(value):
        close = getattr(value, "close", None)
        if close:
            close()
        raise TypeError("worksheet functions must be synchronous; use @xlfly.command for async work")
    return value


def call_functions(batch_json: str) -> str:
    """Evaluate a batch of worksheet calls in function context."""
    results = []
    with running(FUNCTION):
        for item in json.loads(batch_json):
            try:
                results.append({"id": item["id"], "ok": True, "value": to_wire(_invoke(item))})
            except BaseException as exc:
                if isinstance(exc, (KeyboardInterrupt, SystemExit)):
                    exc = RuntimeError(f"function raised {type(exc).__name__}")
                results.append({"id": item["id"], "ok": False, "error": _error(exc)})
    return json.dumps(results)


async def run_command(name: str, token: str) -> str:
    fn = REGISTRY.commands.get(name.upper())
    if fn is None:
        return json.dumps({"ok": False, "error": {"code": "NOT_FOUND", "message": f"no command named {name!r}"}})
    with running(COMMAND, token):
        try:
            value = fn()
            if inspect.isawaitable(value):
                value = await value
            return json.dumps({"ok": True, "value": to_wire(value)})
        except BaseException as exc:
            if isinstance(exc, (KeyboardInterrupt, SystemExit)):
                exc = RuntimeError(f"command raised {type(exc).__name__}")
            return json.dumps({"ok": False, "error": _error(exc, include_trace=True)})


async def dispatch_event(event_json: str, token: str) -> str:
    event = json.loads(event_json)
    errors = []
    with running(COMMAND, token):
        for handler in list(REGISTRY.events.get(event.get("type", ""), [])):
            try:
                value = handler(event)
                if inspect.isawaitable(value):
                    await value
            except BaseException as exc:
                if isinstance(exc, (KeyboardInterrupt, SystemExit)):
                    exc = RuntimeError(f"handler raised {type(exc).__name__}")
                errors.append({"handler": getattr(handler, "__name__", "?"), **_error(exc, include_trace=True)})
    return json.dumps({"ok": not errors, "errors": errors})
````

## 16 python/xlfly/_api.py

The public command API keeps workbook operations bulk and asynchronous. Range.read decodes values; write converts and sizes the matrix; to_frame uses pandas. chart requests a native chart. image converts a figure or PNG bytes to base64. call delegates scalar Excel function calls; log writes to the pane through the guarded host.

Original size: 3622 bytes. SHA-256: `33d4b8fc4b0d0e152bac9cfaa991d800ff937db3ab41c7c461ec2978a25501ff`.

````python
"""Workbook API for commands and event handlers. All calls are bulk and async."""

from __future__ import annotations

import base64
import io
from typing import Any

from . import _host
from ._types import from_wire, to_wire

CHART_KINDS = frozenset({"line", "scatter", "column", "bar", "area", "pie", "histogram", "box"})


class Range:
    """A workbook range, e.g. ``Range("Data!A1:D100")`` or ``Range("A1")`` on the active sheet."""

    def __init__(self, address: str):
        if not isinstance(address, str) or not address.strip():
            raise ValueError("address must be a non-empty string")
        self.address = address.strip()

    def __repr__(self) -> str:
        return f"Range({self.address!r})"

    async def read(self) -> list[list[Any]]:
        """Values as a 2-D list. Errors become XLError, empty cells None."""
        result = await _host.request("range_read", {"address": self.address})
        return from_wire(result["values"])

    async def to_frame(self, header: bool = True):
        """Values as a pandas DataFrame."""
        import pandas as pd

        rows = await self.read()
        if header and rows:
            return pd.DataFrame(rows[1:], columns=[str(c) for c in rows[0]])
        return pd.DataFrame(rows)

    async def write(self, values: Any, header: bool = True) -> str:
        """Write values starting at the range's top-left cell. Returns the written address."""
        wire = to_wire(values, header=header)
        if not isinstance(wire, list):
            wire = [[wire]]
        result = await _host.request("range_write", {"address": self.address, "values": wire})
        return result["address"]

    async def from_frame(self, frame: Any, header: bool = True) -> str:
        return await self.write(frame, header=header)


async def call(name: str, *args: Any) -> Any:
    """Call an Excel worksheet function with values, e.g. ``await call("SUM", 1, 2)``."""
    result = await _host.request("excel_call", {"name": name, "args": [to_wire(a) for a in args]})
    return from_wire(result["value"])


async def chart(kind: str, source: str, anchor: str, title: str | None = None) -> str:
    """Insert a native, editable Excel chart over ``source``. Returns the chart name."""
    if kind not in CHART_KINDS:
        raise ValueError(f"unknown chart kind {kind!r}; expected one of {sorted(CHART_KINDS)}")
    result = await _host.request(
        "chart_native", {"kind": kind, "source": source, "anchor": anchor, "title": title}
    )
    return result["name"]


def _figure_png(figure: Any, dpi: int) -> bytes:
    if isinstance(figure, (bytes, bytearray)):
        return bytes(figure)
    savefig = getattr(figure, "savefig", None)
    if savefig is None:
        raise TypeError("image() expects a matplotlib Figure or PNG bytes")
    buffer = io.BytesIO()
    savefig(buffer, format="png", dpi=dpi, bbox_inches="tight")
    return buffer.getvalue()


async def image(figure: Any, anchor: str, name: str | None = None, dpi: int = 144) -> str:
    """Insert a matplotlib figure (or PNG bytes) as a picture at ``anchor``.

    If a picture with the same ``name`` exists it is replaced.
    """
    png = _figure_png(figure, dpi)
    if not png.startswith(b"\x89PNG"):
        raise ValueError("image data is not a PNG")
    result = await _host.request(
        "image_insert",
        {"base64": base64.b64encode(png).decode("ascii"), "anchor": anchor, "name": name},
    )
    return result["name"]


async def log(message: str) -> None:
    """Write a line to the xlfly task pane log."""
    await _host.request("log", {"message": str(message)})
````

## 17 python/xlfly/__init__.py

Exports the supported public names and package version. Workbook modules should import xlfly and use these APIs rather than relying on internal runtime modules. The opening docstring provides small usage examples.

Original size: 939 bytes. SHA-256: `4b5bcf22cec2a02f19918dbec12eb8b332fe53223664b19f1a9e30cb0d75f035`.

````python
"""xlfly: local Python for Excel.

Worksheet functions::

    import xlfly

    @xlfly.func
    def growth(values, rate):
        return [[v * (1 + rate)] for row in values for v in row]

    # In a cell: =XLFLY.CALL("GROWTH", A1:A10, 0.05)

Commands (run from the task pane, may read and write the workbook)::

    @xlfly.command
    async def report():
        df = await xlfly.Range("Data!A1:C100").to_frame()
        await xlfly.Range("Report!A1").write(df.describe().reset_index())
"""

from ._api import Range, call, chart, image, log
from ._context import ContextError
from ._host import HostError
from ._registry import command, func, on
from ._types import XLError, from_excel_date, to_excel_date

__version__ = "0.1.0"

__all__ = [
    "ContextError",
    "HostError",
    "Range",
    "XLError",
    "call",
    "chart",
    "command",
    "from_excel_date",
    "func",
    "image",
    "log",
    "on",
    "to_excel_date",
]
````

## 18 src/host/excel-services.ts

Implements the Office.js side of the API with load and sync calls. Range reads and writes enforce the cell cap; writes resize from the top-left anchor. Chart insertion maps supported names to Office chart types. Named images attempt replacement. WriteTracker suppresses recent normalized addresses; it is a heuristic that drops sheet identity.

Original size: 5702 bytes. SHA-256: `385584d6f7102e85fa95e0b66ff40ca78b0b81252c92f3201c47f4fafb106c2b`.

````typescript
// Office.js implementations of the operations Python may request.

/* global Excel */

import { RpcError } from "../shared/rpc.ts";
import { isWireError, parseAddress, rangeToWire, wireToRangeValues, type Wire } from "../shared/values.ts";
import type { HostServices } from "./kernel-client.ts";

const CHART_TYPES: Record<string, string> = {
  line: "Line",
  scatter: "XYScatter",
  column: "ColumnClustered",
  bar: "BarClustered",
  area: "Area",
  pie: "Pie",
  histogram: "Histogram",
  box: "Boxwhisker",
};

const MAX_CELLS = 2_000_000;

/** Addresses xlfly itself wrote recently, so the change events they cause can be ignored. */
export class WriteTracker {
  private readonly recent = new Map<string, number>();

  record(address: string, windowMs = 1500): void {
    this.recent.set(normalise(address), Date.now() + windowMs);
  }

  isOwnWrite(address: string): boolean {
    const now = Date.now();
    for (const [key, expires] of this.recent) if (expires < now) this.recent.delete(key);
    return this.recent.has(normalise(address));
  }
}

function normalise(address: string): string {
  return address.replace(/\$/g, "").replace(/^.*!/, "").toUpperCase();
}

function getRange(context: Excel.RequestContext, address: string): Excel.Range {
  const { sheet, range } = parseAddress(address);
  const worksheet = sheet ? context.workbook.worksheets.getItem(sheet) : context.workbook.worksheets.getActiveWorksheet();
  return worksheet.getRange(range);
}

function getSheet(context: Excel.RequestContext, address: string): Excel.Worksheet {
  const { sheet } = parseAddress(address);
  return sheet ? context.workbook.worksheets.getItem(sheet) : context.workbook.worksheets.getActiveWorksheet();
}

export function createExcelServices(tracker: WriteTracker, log: (text: string) => void): HostServices {
  return {
    async range_read({ address }: { address: string }) {
      return Excel.run(async (context) => {
        const range = getRange(context, address);
        range.load(["rowCount", "columnCount"]);
        await context.sync();
        if (range.rowCount * range.columnCount > MAX_CELLS) {
          throw new RpcError("TOO_LARGE", `range has more than ${MAX_CELLS} cells`);
        }
        range.load(["values", "valueTypes"]);
        await context.sync();
        return { values: rangeToWire(range.values, range.valueTypes as unknown as string[][]) };
      });
    },

    async range_write({ address, values }: { address: string; values: Wire }) {
      const rows = wireToRangeValues(values);
      const cells = rows.length * rows[0].length;
      if (cells > MAX_CELLS) throw new RpcError("TOO_LARGE", `write has more than ${MAX_CELLS} cells`);
      return Excel.run(async (context) => {
        const target = getRange(context, address).getCell(0, 0).getResizedRange(rows.length - 1, rows[0].length - 1);
        target.load("address");
        await context.sync();
        tracker.record(target.address);
        target.values = rows;
        await context.sync();
        return { address: target.address };
      });
    },

    async excel_call({ name, args }: { name: string; args: Wire[] }) {
      return Excel.run(async (context) => {
        const functions = context.workbook.functions as any;
        const wanted = String(name).toLowerCase().replace(/[._]/g, "");
        const method = Object.getOwnPropertyNames(Object.getPrototypeOf(functions)).find(
          (key) => key.toLowerCase().replace(/[._]/g, "") === wanted && typeof functions[key] === "function",
        );
        if (!method) throw new RpcError("NOT_SUPPORTED", `Excel function ${name} is not available to add-ins`);
        const scalarArgs = args.map((a) => {
          if (Array.isArray(a) || isWireError(a)) throw new RpcError("BAD_ARGUMENT", "excel_call accepts scalar arguments only");
          return a;
        });
        const result = functions[method](...scalarArgs) as Excel.FunctionResult<unknown>;
        result.load(["value", "error"]);
        await context.sync();
        return { value: result.error ? { $err: result.error } : result.value };
      });
    },

    async chart_native({ kind, source, anchor, title }: { kind: string; source: string; anchor: string; title?: string }) {
      const type = CHART_TYPES[kind];
      if (!type) throw new RpcError("BAD_ARGUMENT", `unknown chart kind ${kind}`);
      return Excel.run(async (context) => {
        const sheet = getSheet(context, anchor);
        const chart = sheet.charts.add(type as Excel.ChartType, getRange(context, source), Excel.ChartSeriesBy.auto);
        chart.setPosition(getRange(context, anchor));
        if (title) chart.title.text = title;
        chart.load("name");
        await context.sync();
        return { name: chart.name };
      });
    },

    async image_insert({ base64, anchor, name }: { base64: string; anchor: string; name?: string }) {
      return Excel.run(async (context) => {
        const sheet = getSheet(context, anchor);
        const cell = getRange(context, anchor).getCell(0, 0);
        cell.load(["left", "top"]);
        if (name) {
          const existing = sheet.shapes.getItemOrNullObject(name);
          await context.sync();
          if (!existing.isNullObject) existing.delete();
        }
        const shape = sheet.shapes.addImage(base64);
        shape.left = cell.left;
        shape.top = cell.top;
        shape.name = name || `xlfly_image_${Date.now()}`;
        shape.altTextDescription = "Created by xlfly";
        shape.load("name");
        await context.sync();
        return { name: shape.name };
      });
    },

    async log({ message }: { message: string }) {
      log(message);
      return null;
    },
  };
}
````

## 19 src/shared/codestore-format.ts

Serializes versioned workbook modules into base64 JSON inside XML. parse validates the document shape. trustKey hashes length-prefixed location, workbook ID, module names and sources sorted by name. Sorting makes module order irrelevant; base64 is encoding, not confidentiality.

Original size: 2370 bytes. SHA-256: `10af4b6d65fd9b53ee36427ea186503bbe23beb575bb8e346cd74bdf54591878`.

````typescript
// Serialisation of workbook code into a custom XML part, and trust hashing.

export const CODE_NAMESPACE = "urn:xlfly:code:v1";

export interface Module {
  name: string;
  source: string;
}

export interface CodeDocument {
  version: 1;
  workbookId: string;
  modules: Module[];
}

function toBase64(text: string): string {
  const bytes = new TextEncoder().encode(text);
  let binary = "";
  for (let i = 0; i < bytes.length; i += 0x8000) {
    binary += String.fromCharCode(...bytes.subarray(i, i + 0x8000));
  }
  return btoa(binary);
}

function fromBase64(encoded: string): string {
  const binary = atob(encoded);
  const bytes = new Uint8Array(binary.length);
  for (let i = 0; i < binary.length; i++) bytes[i] = binary.charCodeAt(i);
  return new TextDecoder().decode(bytes);
}

export function serialize(doc: CodeDocument): string {
  return `<xlfly xmlns="${CODE_NAMESPACE}" version="1">${toBase64(JSON.stringify(doc))}</xlfly>`;
}

export function parse(xml: string): CodeDocument {
  const match = /<xlfly\b[^>]*>([A-Za-z0-9+/=\s]*)<\/xlfly>/.exec(xml);
  if (!match) throw new Error("not an xlfly code part");
  const doc = JSON.parse(fromBase64(match[1].replace(/\s+/g, "")));
  if (doc?.version !== 1 || typeof doc.workbookId !== "string" || !Array.isArray(doc.modules)) {
    throw new Error("unsupported xlfly code part");
  }
  for (const m of doc.modules) {
    if (typeof m?.name !== "string" || typeof m?.source !== "string") throw new Error("malformed module");
  }
  return doc as CodeDocument;
}

export function newWorkbookId(): string {
  return crypto.randomUUID();
}

async function sha256Hex(text: string): Promise<string> {
  const digest = await crypto.subtle.digest("SHA-256", new TextEncoder().encode(text));
  return Array.from(new Uint8Array(digest), (b) => b.toString(16).padStart(2, "0")).join("");
}

/**
 * Trust key: changes whenever the workbook identity, its file location or any module changes.
 * Length-prefixed so that different splits of the same text cannot collide.
 */
export async function trustKey(doc: CodeDocument, documentUrl: string): Promise<string> {
  const parts = [documentUrl, doc.workbookId];
  for (const m of [...doc.modules].sort((a, b) => (a.name < b.name ? -1 : a.name > b.name ? 1 : 0))) {
    parts.push(m.name, m.source);
  }
  return sha256Hex(parts.map((p) => `${p.length}:${p}`).join("|"));
}
````

## 20 src/host/office-store.ts

Loads and saves the workbook custom XML part through Office.js. Save deletes earlier matching parts before adding the new document. Trust timestamps live in OfficeRuntime.storage or localStorage and fail closed on a storage read error. The workbook code and local trust decision are deliberately separate.

Original size: 1924 bytes. SHA-256: `3ff9cc90f530f7374b58e9608cfae99af96ba565040b851edbf7d2bb801c58b0`.

````typescript
// Workbook code storage (custom XML part) and per-user trust decisions.

/* global Excel, OfficeRuntime */

import { CODE_NAMESPACE, newWorkbookId, parse, serialize, type CodeDocument } from "../shared/codestore-format.ts";

export async function loadCode(): Promise<CodeDocument | null> {
  return Excel.run(async (context) => {
    const part = context.workbook.customXmlParts.getByNamespace(CODE_NAMESPACE).getOnlyItemOrNullObject();
    await context.sync();
    if (part.isNullObject) return null;
    const xml = part.getXml();
    await context.sync();
    return parse(xml.value);
  });
}

export async function saveCode(doc: CodeDocument): Promise<void> {
  await Excel.run(async (context) => {
    const parts = context.workbook.customXmlParts.getByNamespace(CODE_NAMESPACE);
    parts.load("items");
    await context.sync();
    for (const part of parts.items) part.delete();
    context.workbook.customXmlParts.add(serialize(doc));
    await context.sync();
  });
}

export function emptyDocument(): CodeDocument {
  return { version: 1, workbookId: newWorkbookId(), modules: [] };
}

const TRUST_PREFIX = "xlfly.trust.";

interface KeyValueStore {
  getItem(key: string): Promise<string | null> | string | null;
  setItem(key: string, value: string): Promise<void> | void;
  removeItem(key: string): Promise<void> | void;
}

function store(): KeyValueStore {
  const runtime = (globalThis as any).OfficeRuntime;
  if (runtime?.storage) return runtime.storage as KeyValueStore;
  return globalThis.localStorage;
}

export async function isTrusted(key: string): Promise<boolean> {
  try {
    return (await store().getItem(TRUST_PREFIX + key)) !== null;
  } catch {
    return false;
  }
}

export async function setTrusted(key: string, trusted: boolean): Promise<void> {
  if (trusted) await store().setItem(TRUST_PREFIX + key, new Date().toISOString());
  else await store().removeItem(TRUST_PREFIX + key);
}
````

## 21 src/app.ts

Wires the system together. Construction creates the kernel client, host services and batcher. init loads code and trust; evaluate gates formulas and maps wire errors to Excel errors. Saving updates code and trust and recalculates. Event registration forwards callbacks after heuristic self-write suppression. Bootstrap associates static custom-function IDs once and binds the pane on Office readiness.

Original size: 9691 bytes. SHA-256: `dfc85ca350cc545137ca780fb1f92f81d72b92a6bf8e6e92ee99cc1a6a40abd9`.

````typescript
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
````

## 22 src/taskpane/taskpane.html

Provides accessible control IDs for the DOM binding and loads Office.js plus the built application. The Python example in its hint uses a scalar cell. Its external Office.js script is why self-hosted Pyodide alone does not imply an offline add-in.

Original size: 1814 bytes. SHA-256: `d4846991207fbe94e3813a86b4b569b8b7dbe241c392d3419e0f65c647fb751b`.

````html
<!doctype html>
<html lang="en">
<head>
  <meta charset="utf-8">
  <meta name="viewport" content="width=device-width, initial-scale=1">
  <title>xlfly</title>
  <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
  <link rel="stylesheet" href="taskpane.css">
</head>
<body>
  <main id="xlfly-root">
    <header>
      <h1>xlfly <small id="version"></small></h1>
      <span id="kernel-status" data-state="stopped">stopped</span>
      <button id="restart-button" title="Restart the Python kernel">Restart</button>
    </header>

    <section id="trust" data-trusted="false">
      <p id="trust-text"></p>
      <button id="trust-button"></button>
    </section>

    <section>
      <div class="row">
        <h2>Modules</h2>
        <select id="module-select" aria-label="Module"></select>
        <button id="new-module">New</button>
        <button id="delete-module">Delete</button>
      </div>
      <form id="new-module-form" class="row" hidden>
        <input id="new-module-name" aria-label="New module name" placeholder="Module name (a Python identifier)" autocomplete="off" spellcheck="false">
        <button type="submit">Add</button>
        <button type="button" id="new-module-cancel">Cancel</button>
      </form>
      <textarea id="editor" spellcheck="false" aria-label="Python source"></textarea>
      <button id="save-button" class="primary">Save to workbook and reload</button>
    </section>

    <section>
      <h2>Functions</h2>
      <p class="hint">Evaluate inline code with <code>=XLFLY.PY("arg1 * 2", A1)</code>.</p>
      <ul id="functions"></ul>
      <h2>Commands</h2>
      <ul id="commands"></ul>
    </section>

    <section>
      <h2>Log</h2>
      <pre id="log"></pre>
    </section>
  </main>
  <script src="app.js"></script>
</body>
</html>
````

## 23 src/taskpane/taskpane.css

Defines light and dark colors, responsive editor width, status styling and readable logs. The hidden attribute is explicitly enforced for the inline New form. This file changes presentation rather than the execution model.

Original size: 2538 bytes. SHA-256: `a0b509136b1f70312620690164ded051f2fcf736009309cdfcc0062ac2ea6dfe`.

````css
:root {
  --bg: #ffffff;
  --fg: #1f1f1f;
  --muted: #616161;
  --line: #e0e0e0;
  --accent: #217346;
  --warn-bg: #fff4ce;
  --ok-bg: #dff6dd;
  --code-bg: #f5f5f5;
  color-scheme: light dark;
}

@media (prefers-color-scheme: dark) {
  :root {
    --bg: #1f1f1f;
    --fg: #f0f0f0;
    --muted: #a0a0a0;
    --line: #3a3a3a;
    --accent: #4caf7a;
    --warn-bg: #433519;
    --ok-bg: #1f3a24;
    --code-bg: #2a2a2a;
  }
}

* { box-sizing: border-box; }

body {
  margin: 0;
  background: var(--bg);
  color: var(--fg);
  font: 13px/1.4 -apple-system, "Segoe UI", system-ui, sans-serif;
}

main { padding: 12px; display: grid; gap: 14px; }

header { display: flex; align-items: center; gap: 8px; }
header h1 { font-size: 16px; margin: 0; flex: 1; }
header small { color: var(--muted); font-weight: normal; }

h2 { font-size: 13px; margin: 0 0 6px; }

#kernel-status {
  font-size: 11px;
  padding: 2px 8px;
  border-radius: 10px;
  background: var(--code-bg);
}
#kernel-status[data-state="ready"] { background: var(--ok-bg); }
#kernel-status[data-state="failed"] { background: var(--warn-bg); }

#trust { padding: 10px; border-radius: 6px; background: var(--warn-bg); }
#trust[data-trusted="true"] { background: var(--ok-bg); }
#trust p { margin: 0 0 8px; }

.row { display: flex; align-items: center; gap: 6px; margin-bottom: 6px; }
.row h2 { margin: 0; flex: 1; }
.row input { flex: 1; min-width: 0; }
[hidden] { display: none !important; }

textarea {
  width: 100%;
  min-height: 240px;
  resize: vertical;
  font: 12px/1.45 ui-monospace, Menlo, Consolas, monospace;
  background: var(--code-bg);
  color: var(--fg);
  border: 1px solid var(--line);
  border-radius: 4px;
  padding: 8px;
  tab-size: 4;
}

button, select, input {
  font: inherit;
  padding: 4px 10px;
  border: 1px solid var(--line);
  border-radius: 4px;
  background: var(--bg);
  color: var(--fg);
}
button.primary { background: var(--accent); border-color: var(--accent); color: #fff; margin-top: 6px; }
button:disabled { opacity: 0.5; }

ul { list-style: none; margin: 0; padding: 0; display: grid; gap: 6px; }
li { display: grid; gap: 2px; }
li span { color: var(--muted); }
li button { justify-self: start; }
code { font: 12px ui-monospace, Menlo, Consolas, monospace; overflow-wrap: anywhere; }
.hint { color: var(--muted); margin: 0 0 6px; }

#log {
  margin: 0;
  max-height: 200px;
  overflow: auto;
  white-space: pre-wrap;
  font: 11px/1.4 ui-monospace, Menlo, Consolas, monospace;
  background: var(--code-bg);
  padding: 8px;
  border-radius: 4px;
}
````

## 24 src/taskpane/taskpane.ts

Maintains draft modules and binds controls to App. Trust changes enable editing; saving persists sources and reloads. New uses an inline form, validates names and avoids window.prompt. Catalog entries use textContent for names and documentation. Status updates preserve dirty drafts; command buttons call App.runCommand.

Original size: 6296 bytes. SHA-256: `5ae9265eb38a308c90e38a9484c9e4096accff7ea77a436cd9c9193bca4c5925`.

````typescript
// Task pane: trust, module editor, catalog of functions and commands, log.

import type { App } from "../app.ts";
import type { Module } from "../shared/codestore-format.ts";

const STARTER = `import xlfly


@xlfly.func
def double(x):
    """=XLFLY.CALL("DOUBLE", A1)"""
    return x * 2


@xlfly.command
async def hello():
    """Writes a greeting to A1 of the active sheet."""
    await xlfly.Range("A1").write("Hello from Python")
`;

export function bindTaskpane(app: App, root: HTMLElement): void {
  const $ = <T extends HTMLElement>(id: string) => root.querySelector<T>(`#${id}`)!;
  const status = $("kernel-status");
  const trustBox = $("trust");
  const trustText = $("trust-text");
  const trustButton = $<HTMLButtonElement>("trust-button");
  const moduleSelect = $<HTMLSelectElement>("module-select");
  const editor = $<HTMLTextAreaElement>("editor");
  const saveButton = $<HTMLButtonElement>("save-button");
  const newButton = $<HTMLButtonElement>("new-module");
  const deleteButton = $<HTMLButtonElement>("delete-module");
  const newForm = $<HTMLFormElement>("new-module-form");
  const newName = $<HTMLInputElement>("new-module-name");
  const functions = $("functions");
  const commands = $("commands");
  const log = $("log");
  $("version").textContent = `v${app.version}`;

  let drafts: Module[] = [];
  let selected = 0;
  let dirty = false;

  const loadDrafts = () => {
    drafts = (app.doc?.modules ?? []).map((m) => ({ ...m }));
    if (drafts.length === 0) drafts.push({ name: "main", source: STARTER });
    selected = Math.min(selected, drafts.length - 1);
    dirty = false;
  };

  const renderModules = () => {
    moduleSelect.replaceChildren(
      ...drafts.map((m, i) => {
        const option = document.createElement("option");
        option.value = String(i);
        option.textContent = m.name;
        option.selected = i === selected;
        return option;
      }),
    );
    editor.value = drafts[selected]?.source ?? "";
    editor.readOnly = !app.trusted;
  };

  const render = () => {
    status.textContent = app.kernelStatus;
    status.dataset.state = app.kernelStatus;
    status.title = app.kernelDetail;

    trustBox.dataset.trusted = String(app.trusted);
    trustText.textContent = app.trusted
      ? "Python code in this workbook is trusted on this computer."
      : "Python code in this workbook is not trusted and will not run. Review the code below, then trust it.";
    trustButton.textContent = app.trusted ? "Revoke trust" : "Trust this workbook";
    saveButton.disabled = !app.trusted;
    newButton.disabled = !app.trusted;
    deleteButton.disabled = !app.trusted || drafts.length <= 1;
    editor.readOnly = !app.trusted;

    const catalog = app.kernel.catalog;
    functions.replaceChildren(
      ...Object.entries(catalog.functions).map(([name, info]) => {
        const li = document.createElement("li");
        li.innerHTML = `<code></code><span></span>`;
        li.querySelector("code")!.textContent = `=XLFLY.CALL("${name}"${info.signature === "()" ? "" : ", …"})`;
        li.querySelector("span")!.textContent = info.doc;
        return li;
      }),
    );
    commands.replaceChildren(
      ...Object.entries(catalog.commands).map(([name, info]) => {
        const li = document.createElement("li");
        const button = document.createElement("button");
        button.textContent = `Run ${name}`;
        button.onclick = () => void app.runCommand(name);
        const doc = document.createElement("span");
        doc.textContent = info.doc;
        li.append(button, doc);
        return li;
      }),
    );
    log.textContent = app.logLines.join("\n");
    log.scrollTop = log.scrollHeight;
  };

  moduleSelect.onchange = () => {
    drafts[selected].source = editor.value;
    selected = Number(moduleSelect.value);
    renderModules();
  };
  editor.oninput = () => {
    dirty = true;
  };
  editor.onkeydown = (event) => {
    if (event.key === "Tab") {
      event.preventDefault();
      editor.setRangeText("    ", editor.selectionStart, editor.selectionEnd, "end");
      dirty = true;
    }
  };
  // window.prompt() is unavailable in Office add-in web views, so the name is asked for inline.
  newButton.onclick = () => {
    newForm.hidden = false;
    newName.value = `module${drafts.length + 1}`;
    newName.focus();
    newName.select();
  };
  $("new-module-cancel").onclick = () => {
    newForm.hidden = true;
  };
  newName.onkeydown = (event) => {
    if (event.key === "Escape") newForm.hidden = true;
  };
  newForm.onsubmit = (event) => {
    event.preventDefault();
    const name = newName.value.trim();
    if (!/^[A-Za-z_]\w*$/.test(name) || drafts.some((m) => m.name === name)) {
      app.log(`invalid or duplicate module name: ${name}`);
      newName.focus();
      return;
    }
    newForm.hidden = true;
    drafts[selected].source = editor.value;
    drafts.push({ name, source: "import xlfly\n" });
    selected = drafts.length - 1;
    dirty = true;
    renderModules();
    render();
  };
  deleteButton.onclick = () => {
    if (drafts.length <= 1) return;
    drafts.splice(selected, 1);
    selected = 0;
    dirty = true;
    renderModules();
    render();
  };
  saveButton.onclick = async () => {
    drafts[selected].source = editor.value;
    saveButton.disabled = true;
    try {
      await app.saveModules(drafts.map((m) => ({ ...m })));
      dirty = false;
    } catch (error) {
      app.log(`save failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      saveButton.disabled = !app.trusted;
    }
  };
  trustButton.onclick = async () => {
    trustButton.disabled = true;
    try {
      if (app.trusted) await app.revokeTrust();
      else await app.trustWorkbook();
    } catch (error) {
      app.log(`trust change failed: ${error instanceof Error ? error.message : String(error)}`);
    } finally {
      trustButton.disabled = false;
    }
  };
  $("restart-button").onclick = () => {
    app.kernel.restart("restarted by user");
    void app.reload();
  };

  let lastDoc = app.doc;
  app.subscribe(() => {
    if (app.doc !== lastDoc && !dirty) {
      lastDoc = app.doc;
      loadDrafts();
      renderModules();
    }
    render();
  });

  loadDrafts();
  renderModules();
  render();
}
````

## 25 build.mjs

Cleans and rebuilds dist, bundles the host as IIFE and the worker as ESM, copies UI and metadata, serializes Python package files and copies the Pyodide runtime. Optional vendor wheels are copied into the runtime directory. Environment variables select deployment URLs; they are build inputs rather than workbook settings.

Original size: 2542 bytes. SHA-256: `6e71f5f2c65f87a84d0e7e49ebedfe396aa838629a8b70debe82bb1250a3ff66`.

````javascript
// Build the add-in into dist/: bundles, static files, Python bundle, Pyodide runtime, manifest.
import * as esbuild from "esbuild";
import { cpSync, existsSync, mkdirSync, readdirSync, readFileSync, rmSync, statSync, writeFileSync } from "node:fs";
import { join, relative } from "node:path";

const pkg = JSON.parse(readFileSync("package.json", "utf8"));
const baseUrl = (process.env.XLFLY_BASE_URL ?? "https://localhost:3000").replace(/\/+$/, "");
const packageBaseUrl = process.env.XLFLY_PACKAGE_BASE_URL ?? "";
const out = "dist";

rmSync(out, { recursive: true, force: true });
mkdirSync(join(out, "pyodide"), { recursive: true });

const common = {
  bundle: true,
  target: ["es2020", "safari15"],
  sourcemap: true,
  logLevel: "warning",
  define: {
    XLFLY_VERSION: JSON.stringify(pkg.version),
    XLFLY_PACKAGE_BASE_URL: JSON.stringify(packageBaseUrl),
  },
};
await esbuild.build({ ...common, entryPoints: ["src/app.ts"], outfile: join(out, "app.js"), format: "iife" });
await esbuild.build({ ...common, entryPoints: ["src/kernel/worker.ts"], outfile: join(out, "kernel.js"), format: "esm" });

cpSync("src/taskpane/taskpane.html", join(out, "taskpane.html"));
cpSync("src/taskpane/taskpane.css", join(out, "taskpane.css"));
cpSync("src/functions/functions.json", join(out, "functions.json"));
cpSync("assets", join(out, "assets"), { recursive: true });

// Python package -> one JSON file the worker writes into the Pyodide file system.
const python = {};
const walk = (dir) => {
  for (const name of readdirSync(dir)) {
    const path = join(dir, name);
    if (statSync(path).isDirectory()) { if (name !== "__pycache__") walk(path); }
    else if (name.endsWith(".py")) python[relative("python", path)] = readFileSync(path, "utf8");
  }
};
walk("python/xlfly");
writeFileSync(join(out, "python-bundle.json"), JSON.stringify(python));

// Pyodide core runtime from node_modules, plus any packages fetched into vendor/pyodide.
for (const file of ["pyodide.mjs", "pyodide.asm.mjs", "pyodide.asm.wasm", "python_stdlib.zip", "pyodide-lock.json"]) {
  cpSync(join("node_modules/pyodide", file), join(out, "pyodide", file));
}
if (existsSync("vendor/pyodide")) cpSync("vendor/pyodide", join(out, "pyodide"), { recursive: true });
else if (!packageBaseUrl) console.warn("note: vendor/pyodide is empty; run `npm run fetch:pyodide` so numpy/pandas/matplotlib can load");

writeFileSync(join(out, "manifest.xml"), readFileSync("manifest.xml", "utf8").replaceAll("{{BASE_URL}}", baseUrl));
console.log(`built ${out}/ for ${baseUrl}`);
````

## 26 scripts/fetch-pyodide-packages.mjs

Computes the transitive package dependency closure from the installed Pyodide lockfile. Each existing or downloaded file is SHA-256 checked before use. The default package roots are numpy, pandas and matplotlib. This is a setup-time download, not installation of arbitrary packages from a cell.

Original size: 1922 bytes. SHA-256: `879a087472fdc7f35803c7232bc7ad939cca13c9e0cab410cb99e55cb4afff88`.

````javascript
// Download numpy, pandas, matplotlib and their dependencies from the Pyodide release
// into vendor/pyodide, verifying each file against the SHA-256 in pyodide-lock.json.
// Self-hosting keeps the add-in independent of third-party CDNs at run time.
import { createHash } from "node:crypto";
import { existsSync, mkdirSync, readFileSync, writeFileSync } from "node:fs";
import { join } from "node:path";

const lock = JSON.parse(readFileSync("node_modules/pyodide/pyodide-lock.json", "utf8"));
const version = JSON.parse(readFileSync("node_modules/pyodide/package.json", "utf8")).version;
const source = (process.env.XLFLY_PYODIDE_SOURCE ?? `https://cdn.jsdelivr.net/pyodide/v${version}/full/`).replace(/\/?$/, "/");
const wanted = (process.env.XLFLY_PACKAGES ?? "numpy,pandas,matplotlib").split(",").map((s) => s.trim().toLowerCase());
const outDir = "vendor/pyodide";
mkdirSync(outDir, { recursive: true });

const closure = new Set();
const add = (name) => {
  const key = name.toLowerCase();
  if (closure.has(key)) return;
  const entry = lock.packages[key];
  if (!entry) throw new Error(`package ${name} is not in the Pyodide lock file`);
  closure.add(key);
  for (const dep of entry.depends ?? []) add(dep);
};
wanted.forEach(add);

for (const name of closure) {
  const { file_name: file, sha256 } = lock.packages[name];
  const target = join(outDir, file);
  if (existsSync(target) && createHash("sha256").update(readFileSync(target)).digest("hex") === sha256) continue;
  const response = await fetch(source + file);
  if (!response.ok) throw new Error(`${file}: HTTP ${response.status}`);
  const bytes = Buffer.from(await response.arrayBuffer());
  const digest = createHash("sha256").update(bytes).digest("hex");
  if (digest !== sha256) throw new Error(`${file}: checksum mismatch`);
  writeFileSync(target, bytes);
  console.log(`fetched ${file}`);
}
console.log(`${closure.size} packages in ${outDir}`);
````

## 27 scripts/serve.mjs

Serves dist over HTTPS using the local Office development certificate. MIME types include WASM and Python wheels. It disables caching for development and enables CORS. This development server is not presented as production hosting; review production routing, headers and deployment separately.

Original size: 1751 bytes. SHA-256: `8fe988cb6d4746c6afa1268fefe61af54a46a3057c6b6cbdee62b6e62c7f066a`.

````javascript
// HTTPS static server for dist/ during development. Certificates come from office-addin-dev-certs.
import { createServer } from "node:https";
import { createReadStream, existsSync, readFileSync, statSync } from "node:fs";
import { homedir } from "node:os";
import { extname, join, normalize, resolve } from "node:path";

const root = resolve("dist");
const port = Number(process.env.PORT ?? 3000);
const certDir = join(homedir(), ".office-addin-dev-certs");
if (!existsSync(join(certDir, "localhost.crt"))) {
  console.error("No development certificate. Run: npx office-addin-dev-certs install");
  process.exit(1);
}
const types = {
  ".html": "text/html; charset=utf-8", ".js": "text/javascript", ".mjs": "text/javascript",
  ".css": "text/css", ".json": "application/json", ".wasm": "application/wasm",
  ".zip": "application/zip", ".whl": "application/zip", ".png": "image/png",
  ".xml": "application/xml", ".map": "application/json", ".tar": "application/x-tar",
};

createServer(
  { key: readFileSync(join(certDir, "localhost.key")), cert: readFileSync(join(certDir, "localhost.crt")) },
  (req, res) => {
    const path = normalize(decodeURIComponent(new URL(req.url, "https://localhost").pathname));
    const file = join(root, path.endsWith("/") ? join(path, "taskpane.html") : path);
    if (!file.startsWith(root) || !existsSync(file) || !statSync(file).isFile()) {
      res.writeHead(404).end("not found");
      return;
    }
    res.writeHead(200, {
      "Content-Type": types[extname(file)] ?? "application/octet-stream",
      "Cache-Control": "no-store",
      "Access-Control-Allow-Origin": "*",
    });
    createReadStream(file).pipe(res);
  },
).listen(port, () => console.log(`serving dist/ at https://localhost:${port}`));
````

## 28 scripts/setup-mac.sh

Orchestrates the Mac path: verify host prerequisites, npm ci, fetch packages, trust a localhost certificate, build and sideload. It finishes with manual server and Excel restart instructions. The certificate step may require user interaction.

Original size: 1242 bytes. SHA-256: `95bbd5a8cfe54f922d40a2669eaecad69207fb6622bdc85a39557f6980f044c2`.

````bash
#!/usr/bin/env bash
# One-step developer setup on macOS: dependencies, Python packages, HTTPS certificate,
# build and sideload. Afterwards run `npm run serve` and restart Excel.
set -euo pipefail
cd "$(dirname "$0")/.."

[[ "$(uname -s)" == "Darwin" ]] || { echo "This script is for macOS."; exit 1; }
command -v node >/dev/null || { echo "Node.js 22 or later is required (brew install node)."; exit 1; }
major="$(node -p 'process.versions.node.split(".")[0]')"
(( major >= 22 )) || { echo "Node.js 22 or later is required; found $(node -v)."; exit 1; }
[[ -d "/Applications/Microsoft Excel.app" ]] || echo "warning: Microsoft Excel.app not found in /Applications"

echo "==> installing dependencies"
npm ci
echo "==> fetching numpy, pandas, matplotlib for Pyodide (checksums verified)"
npm run fetch:pyodide
echo "==> creating and trusting a localhost HTTPS certificate (macOS may ask for your password)"
npx office-addin-dev-certs install
echo "==> building"
npm run build
echo "==> sideloading the manifest into Excel"
bash scripts/sideload-mac.sh

cat <<'DONE'

Done. Next:
  1. npm run serve          (keep it running)
  2. Quit and reopen Excel, open a workbook
  3. Home > Add-ins > xlfly, or the "Python" button on the Home tab
DONE
````

## 29 scripts/sideload-mac.sh

Creates Excel for Mac’s wef folder and copies the generated manifest into it. The manifest still points to the development HTTPS server, which must stay running. This script is platform-specific even though the application architecture targets multiple hosts.

Original size: 458 bytes. SHA-256: `2a08a74d8f41ddb64c8047bd7e1c3e9f4ce359568c137c992e247b6bf2db3f89`.

````bash
#!/usr/bin/env bash
# Sideload the development manifest into Excel for Mac (Microsoft's documented method:
# place the manifest in Excel's "wef" folder, then restart Excel and use Home > Add-ins).
set -euo pipefail
wef="$HOME/Library/Containers/com.microsoft.Excel/Data/Documents/wef"
mkdir -p "$wef"
cp "$(dirname "$0")/../dist/manifest.xml" "$wef/xlfly.manifest.xml"
echo "Installed $wef/xlfly.manifest.xml — restart Excel, then Home > Add-ins > xlfly."
````

## 30 tests/rpc.test.ts

Pairs two in-memory endpoints and tests requests in both directions, nested callbacks, error codes, timeouts, late replies, shutdown and version mismatch. These tests establish protocol behavior independently of Office or Python.

Original size: 2526 bytes. SHA-256: `00995c8b64f9aaf55b7e326a0a93e9b613334ad61581733b20ae20c4cac67a4d`.

````typescript
import { test } from "node:test";
import assert from "node:assert/strict";
import { RpcEndpoint, RpcError, type Message } from "../src/shared/rpc.ts";

function pair(): [RpcEndpoint, RpcEndpoint] {
  let a: RpcEndpoint, b: RpcEndpoint;
  a = new RpcEndpoint((m: Message) => queueMicrotask(() => void b.handle(structuredClone(m))));
  b = new RpcEndpoint((m: Message) => queueMicrotask(() => void a.handle(structuredClone(m))));
  return [a, b];
}

test("requests flow in both directions", async () => {
  const [host, kernel] = pair();
  kernel.on("add", ({ x, y }) => x + y);
  host.on("ping", (_p, meta) => `pong:${meta.context}`);
  assert.equal(await host.request("add", { x: 2, y: 3 }), 5);
  assert.equal(await kernel.request("ping", null, { context: "tok" }), "pong:tok");
});

test("a handler can call back across the channel while handling a request", async () => {
  const [host, kernel] = pair();
  host.on("read", () => [[1, 2]]);
  kernel.on("run", async () => (await kernel.request<number[][]>("read", {}))[0][1]);
  assert.equal(await host.request("run", {}), 2);
});

test("errors keep their code", async () => {
  const [host, kernel] = pair();
  kernel.on("boom", () => { throw new RpcError("CONTEXT", "not allowed"); });
  await assert.rejects(host.request("boom", {}), (e: any) => e.code === "CONTEXT" && e.message === "CONTEXT: not allowed");
  await assert.rejects(host.request("missing", {}), (e: any) => e.code === "METHOD_NOT_FOUND");
});

test("timeouts reject and late replies are ignored", async () => {
  const [host, kernel] = pair();
  kernel.on("slow", () => new Promise((r) => setTimeout(() => r(1), 50)));
  await assert.rejects(host.request("slow", {}, { timeoutMs: 5 }), (e: any) => e.code === "TIMEOUT");
  await new Promise((r) => setTimeout(r, 80));
  assert.equal(host.outstanding, 0);
});

test("close rejects outstanding requests", async () => {
  const [host, kernel] = pair();
  kernel.on("never", () => new Promise(() => {}));
  const pending = host.request("never", {});
  host.close("kernel died");
  await assert.rejects(pending, (e: any) => e.code === "CLOSED");
  await assert.rejects(host.request("x", {}), (e: any) => e.code === "CLOSED");
});

test("protocol version mismatch is refused", async () => {
  const sent: Message[] = [];
  const endpoint = new RpcEndpoint((m) => sent.push(m));
  await endpoint.handle({ v: 99, kind: "req", id: 1, method: "x", params: null } as any);
  assert.equal(sent.length, 1);
  assert.equal((sent[0] as any).error.code, "PROTOCOL");
});
````

## 31 tests/values.test.ts

Checks scalar and matrix conversion, typed Excel errors, rectangular writes, result-error preservation and sheet-address parsing. These are contract tests for the boundary where subtle spreadsheet behavior becomes code.

Original size: 1497 bytes. SHA-256: `ac1e8236b9ca6c2be94db57beceedcad0c50a91166a328011f14cc3386875267`.

````typescript
import { test } from "node:test";
import assert from "node:assert/strict";
import { argToWire, parseAddress, rangeToWire, wireToMatrix, wireToRangeValues } from "../src/shared/values.ts";

test("custom function arguments", () => {
  assert.equal(argToWire(5), 5);
  assert.equal(argToWire(""), null);
  assert.equal(argToWire(undefined), null);
  assert.deepEqual(argToWire([[1, "", true]]), [[1, null, true]]);
  assert.deepEqual(argToWire({ code: "#N/A" }), { $err: "#N/A" });
});

test("range values with types", () => {
  assert.deepEqual(
    rangeToWire([[1, "#DIV/0!", ""], ["x", false, "#N/A"]], [["Double", "Error", "Empty"], ["String", "Boolean", "String"]]),
    [[1, { $err: "#DIV/0!" }, null], ["x", false, "#N/A"]],
  );
});

test("writes are rectangular and errors become text", () => {
  assert.deepEqual(wireToRangeValues(7), [[7]]);
  assert.deepEqual(wireToRangeValues([[1], [2, { $err: "#N/A" }], [null]]), [[1, ""], [2, "#N/A"], ["", ""]]);
});

test("function results keep errors separate", () => {
  assert.deepEqual(wireToMatrix("a"), [["a"]]);
  assert.deepEqual(wireToMatrix([[1, { $err: "#NUM!" }]]), [[1, { error: { code: "#NUM!", message: "" } }]]);
});

test("addresses", () => {
  assert.deepEqual(parseAddress("A1:B2"), { sheet: null, range: "A1:B2" });
  assert.deepEqual(parseAddress("Data!A1"), { sheet: "Data", range: "A1" });
  assert.deepEqual(parseAddress("'Q1 ''24'!C3"), { sheet: "Q1 '24", range: "C3" });
  assert.throws(() => parseAddress("Data!"));
});
````

## 32 tests/codestore.test.ts

Tests XML-hostile and non-ASCII source round trips, malformed documents and trust hashing. The assertions show that identity, location and source change the key while module order does not. A length-prefix test prevents ambiguous concatenations.

Original size: 1603 bytes. SHA-256: `e16f585cb083593dd1de6a3efafdaca0e8b9cb0c1fcd3bcf7345b3fb46036671`.

````typescript
import { test } from "node:test";
import assert from "node:assert/strict";
import { parse, serialize, trustKey, type CodeDocument } from "../src/shared/codestore-format.ts";

const doc: CodeDocument = {
  version: 1,
  workbookId: "wb-1",
  modules: [{ name: "main", source: "x = ']]>' + 'ünïcode ✓'\n" }],
};

test("round trip survives XML-hostile and non-ASCII source", () => {
  assert.deepEqual(parse(serialize(doc)), doc);
});

test("malformed parts are rejected", () => {
  assert.throws(() => parse("<other/>"));
  assert.throws(() => parse(`<xlfly>${btoa(JSON.stringify({ version: 2 }))}</xlfly>`));
});

test("trust key changes with code, identity and location, not module order", async () => {
  const base = await trustKey(doc, "file:///a.xlsx");
  assert.equal(base, await trustKey(structuredClone(doc), "file:///a.xlsx"));
  assert.notEqual(base, await trustKey(doc, "file:///b.xlsx"));
  assert.notEqual(base, await trustKey({ ...doc, workbookId: "wb-2" }, "file:///a.xlsx"));
  assert.notEqual(base, await trustKey({ ...doc, modules: [{ name: "main", source: "y = 1" }] }, "file:///a.xlsx"));
  const two = { ...doc, modules: [{ name: "a", source: "1" }, { name: "b", source: "2" }] };
  const swapped = { ...doc, modules: [two.modules[1], two.modules[0]] };
  assert.equal(await trustKey(two, ""), await trustKey(swapped, ""));
  // length prefixing: moving text between name and source changes the key
  assert.notEqual(
    await trustKey({ ...doc, modules: [{ name: "ab", source: "c" }] }, ""),
    await trustKey({ ...doc, modules: [{ name: "a", source: "bc" }] }, ""),
  );
});
````

## 33 tests/kernel-client.test.ts

Uses a scripted fake worker to test command permission, stale tokens, timeout termination and module reload, serialized commands and batching. Observe that the timeout test calls ensureStarted after the failure: recovery is on demand.

Original size: 4731 bytes. SHA-256: `374653a31b61e9ae4730245242e01ed1277ae31b79a7f3b349fcd69945114ea5`.

````typescript
import { test } from "node:test";
import assert from "node:assert/strict";
import { RpcEndpoint, type Message } from "../src/shared/rpc.ts";
import { KernelClient, type HostServices, type WorkerLike } from "../src/host/kernel-client.ts";
import { CallBatcher } from "../src/host/batcher.ts";

/** An in-process fake worker whose kernel side is scripted by the test. */
function fakeWorkerFactory(setup: (rpc: RpcEndpoint) => void) {
  const created: { terminated: boolean }[] = [];
  const factory = (): WorkerLike => {
    const state = { terminated: false };
    created.push(state);
    const worker: WorkerLike = {
      onmessage: null,
      onerror: null,
      postMessage: (m) => queueMicrotask(() => { if (!state.terminated) void kernel.handle(structuredClone(m) as Message); }),
      terminate: () => { state.terminated = true; },
    };
    const kernel = new RpcEndpoint((m) => queueMicrotask(() => { if (!state.terminated) worker.onmessage?.({ data: structuredClone(m) }); }));
    kernel.on("start", () => ({ python: "test" }));
    kernel.on("load_modules", () => ({ functions: {}, commands: {}, events: [], errors: [] }));
    setup(kernel);
    return worker;
  };
  return { factory, created };
}

const calls: string[] = [];
const services: HostServices = {
  range_read: async () => { calls.push("read"); return { values: [[1]] }; },
  range_write: async () => ({ address: "A1" }),
  excel_call: async () => ({ value: 1 }),
  chart_native: async () => ({ name: "c" }),
  image_insert: async () => ({ name: "i" }),
  log: async () => null,
};

test("Excel access is refused without a command token and allowed with one", async () => {
  const { factory } = fakeWorkerFactory((rpc) => {
    rpc.on("call_functions", async () => {
      await rpc.request("range_read", { address: "A1" });  // no context: must be refused
      return [];
    });
    rpc.on("run_command", async (_p, meta) => rpc.request("range_read", { address: "A1" }, { context: meta.context }));
  });
  const client = new KernelClient({ createWorker: factory, startParams: {}, services });
  await assert.rejects(client.callFunctions([]), (e: any) => e.code === "CONTEXT");
  calls.length = 0;
  assert.deepEqual(await client.runCommand("x"), { values: [[1]] });
  assert.deepEqual(calls, ["read"]);
});

test("a stale token from a finished command is refused", async () => {
  let saved = "";
  const { factory } = fakeWorkerFactory((rpc) => {
    rpc.on("run_command", (_p, meta) => { saved = meta.context!; return { ok: true }; });
    rpc.on("call_functions", async () => rpc.request("range_read", {}, { context: saved }));
  });
  const client = new KernelClient({ createWorker: factory, startParams: {}, services });
  await client.runCommand("x");
  await assert.rejects(client.callFunctions([]), (e: any) => e.code === "CONTEXT");
});

test("a stuck worksheet function restarts the kernel and modules are reloaded", async () => {
  let loads = 0;
  const { factory, created } = fakeWorkerFactory((rpc) => {
    rpc.on("load_modules", () => { loads++; return { functions: {}, commands: {}, events: [], errors: [] }; });
    rpc.on("call_functions", () => new Promise(() => {}));
  });
  const client = new KernelClient({ createWorker: factory, startParams: {}, services, functionTimeoutMs: 20 });
  await client.loadModules([{ name: "m", source: "" }]);
  await assert.rejects(client.callFunctions([]), (e: any) => e.code === "TIMEOUT");
  assert.equal(created[0].terminated, true);
  await client.ensureStarted();
  assert.equal(created.length, 2);
  assert.equal(loads, 2);
});

test("commands run one at a time", async () => {
  let active = 0, peak = 0;
  const { factory } = fakeWorkerFactory((rpc) => {
    rpc.on("run_command", async () => {
      active++; peak = Math.max(peak, active);
      await new Promise((r) => setTimeout(r, 5));
      active--;
      return { ok: true };
    });
  });
  const client = new KernelClient({ createWorker: factory, startParams: {}, services });
  await Promise.all([client.runCommand("a"), client.runCommand("b"), client.runCommand("c")]);
  assert.equal(peak, 1);
});

test("calls in one tick are sent as one batch", async () => {
  const batches: number[] = [];
  const batcher = new CallBatcher(async (batch) => {
    batches.push(batch.length);
    return batch.map((b) => ({ id: b.id, ok: true as const, value: b.kind }));
  });
  const results = await Promise.all([
    batcher.enqueue({ kind: "eval", code: "1", args: [] }),
    batcher.enqueue({ kind: "call", name: "F", args: [] }),
    batcher.enqueue({ kind: "eval", code: "2", args: [] }),
  ]);
  assert.deepEqual(batches, [3]);
  assert.deepEqual(results.map((r) => (r.ok ? r.value : null)), ["eval", "call", "eval"]);
});
````

## 34 tests/kernel.integration.ts

Runs real Pyodide and the real kernel/client protocol with fake workbook services. It exercises a function, inline evaluation, a forbidden workbook call, a command that reads and writes and charts, stdout and events. It does not test actual Excel APIs or worker termination because its in-process worker has a no-op terminate.

Original size: 4125 bytes. SHA-256: `fc0d0453053ad19db1ed8c87958da401e5d54e9486ddc0cf068ee82cfd6b372a`.

````typescript
// Runs the real Pyodide kernel in-process through KernelClient, with Excel replaced by fakes.
import { test } from "node:test";
import assert from "node:assert/strict";
import { readdirSync, readFileSync } from "node:fs";
import { createRequire } from "node:module";
import { dirname, join } from "node:path";
import { loadPyodide } from "pyodide";
import { RpcEndpoint, type Message } from "../src/shared/rpc.ts";
import { startKernel } from "../src/kernel/kernel-core.ts";
import { KernelClient, type HostServices, type WorkerLike } from "../src/host/kernel-client.ts";

const indexURL = dirname(createRequire(import.meta.url).resolve("pyodide/package.json")) + "/";
const python: Record<string, string> = {};
for (const f of readdirSync("python/xlfly")) if (f.endsWith(".py")) python[`xlfly/${f}`] = readFileSync(join("python/xlfly", f), "utf8");

function inProcessWorker(): WorkerLike {
  const worker: WorkerLike = {
    onmessage: null,
    onerror: null,
    postMessage: (m) => queueMicrotask(() => void kernel.handle(structuredClone(m) as Message)),
    terminate: () => {},
  };
  const kernel = new RpcEndpoint((m) => queueMicrotask(() => worker.onmessage?.({ data: structuredClone(m) })));
  kernel.on("start", () => startKernel(kernel, loadPyodide as any, { indexURL, python }));
  return worker;
}

const sheet = new Map<string, unknown>([["A1", 10], ["A2", 20]]);
const writes: unknown[] = [];
const services: HostServices = {
  range_read: async ({ address }) => ({ values: [[sheet.get("A1")], [sheet.get("A2")]], address }),
  range_write: async ({ address, values }) => { writes.push({ address, values }); return { address: "B1:B2" }; },
  excel_call: async ({ name, args }) => ({ value: name === "SUM" ? (args as number[]).reduce((a, b) => a + b, 0) : null }),
  chart_native: async ({ kind }) => ({ name: `Chart ${kind}` }),
  image_insert: async ({ name }) => ({ name }),
  log: async () => null,
};

const SOURCE = `
import xlfly

@xlfly.func
def scale(values, k):
    return [[v * k] for row in values for v in row]

@xlfly.func
def peek():
    coroutine = xlfly.Range("A1").read()
    coroutine.send(None)  # the first step raises ContextError before any host request

@xlfly.command
async def report():
    values = await xlfly.Range("Data!A1:A2").read()
    total = await xlfly.call("SUM", *[r[0] for r in values])
    await xlfly.Range("B1").write([[v[0] * 2] for v in values])
    name = await xlfly.chart("column", "Data!A1:A2", "Data!D2", title="Values")
    print("report done", total)
    return [total, name]

@xlfly.on("sheet_change")
async def changed(event):
    await xlfly.log("changed " + event["address"])
`;

test("Pyodide kernel end to end", { timeout: 120_000 }, async () => {
  const output: string[] = [];
  const client = new KernelClient({
    createWorker: inProcessWorker,
    startParams: {},
    services,
    onOutput: (_s, text) => output.push(text),
  });

  const catalog = await client.loadModules([{ name: "main", source: SOURCE }]);
  assert.deepEqual(catalog.errors, []);
  assert.deepEqual(Object.keys(catalog.functions), ["PEEK", "SCALE"]);
  assert.match(client.info!.python, /^3\./);

  const results = await client.callFunctions([
    { id: "1", kind: "call", name: "scale", args: [[[1], [2]], 3] },
    { id: "2", kind: "eval", code: "sum(r[0] for r in arg1)", args: [[[1], [2], [3]]] },
    { id: "3", kind: "call", name: "peek", args: [] },
  ]);
  assert.deepEqual(results[0], { id: "1", ok: true, value: [[3], [6]] });
  assert.deepEqual(results[1], { id: "2", ok: true, value: 6 });
  assert.equal(results[2].ok, false);
  assert.match((results[2] as any).error.message, /ContextError/);
  assert.equal(writes.length, 0);

  const command = await client.runCommand("report");
  assert.deepEqual(command, { ok: true, value: [[30], ["Chart column"]] });
  assert.deepEqual(writes, [{ address: "B1", values: [[20], [40]] }]);
  assert.ok(output.some((line) => line.includes("report done 30")));

  const event = await client.dispatchEvent({ type: "sheet_change", address: "C3" });
  assert.deepEqual(event, { ok: true, errors: [] });

  client.stop();
});
````

## 35 python/tests/test_runtime.py

Loads sample workbook modules and tests registration, reload, formula results, exception mapping, context restrictions, async refusal, command tokens and events. The fake host records requests so tests can assert that prohibited calls never reached it.

Original size: 4902 bytes. SHA-256: `fae04b1856e4ab4fcd618b498f2bada9490deeaacc4b6097edc2e71c0ae79aa3`.

````python
import asyncio
import json
import unittest

from xlfly import _host, _runtime

MODULE = '''
import xlfly

@xlfly.func
def double(x):
    """Doubles x."""
    return x * 2

@xlfly.func("TOTAL")
def total(rows):
    return sum(v for row in rows for v in row if isinstance(v, (int, float)))

@xlfly.func
def sneaky():
    import asyncio
    return asyncio.run(xlfly.Range("A1").read())

@xlfly.func
def not_sync():
    async def inner():
        return 1
    return inner()

@xlfly.command
async def copy():
    values = await xlfly.Range("A1:A2").read()
    return await xlfly.Range("B1").write(values)

@xlfly.on("sheet_change")
def changed(event):
    seen.append(event["address"])

seen = []
'''


def load(source=MODULE):
    return json.loads(_runtime.load_modules(json.dumps([{"name": "main", "source": source}])))


def calls(*items):
    batch = [dict(item, id=str(i)) for i, item in enumerate(items)]
    return json.loads(_runtime.call_functions(json.dumps(batch)))


class RuntimeTests(unittest.TestCase):
    def setUp(self):
        self.requests = []

        async def bridge(method, params_json, token):
            self.requests.append((method, json.loads(params_json), token))
            if method == "range_read":
                return json.dumps({"values": [[1], [2]]})
            return json.dumps({"address": "B1:B2"})

        _host.install_bridge(bridge)

    def tearDown(self):
        _host.install_bridge(None)

    def test_catalog(self):
        catalog = load()
        self.assertEqual(sorted(catalog["functions"]), ["DOUBLE", "NOT_SYNC", "SNEAKY", "TOTAL"])
        self.assertEqual(catalog["functions"]["DOUBLE"]["doc"], "Doubles x.")
        self.assertEqual(list(catalog["commands"]), ["COPY"])
        self.assertEqual(catalog["events"], ["sheet_change"])
        self.assertEqual(catalog["errors"], [])

    def test_module_errors_are_reported_not_raised(self):
        catalog = load("import xlfly\n1/0\n")
        self.assertEqual(catalog["errors"][0]["module"], "main")
        self.assertIn("ZeroDivisionError", catalog["errors"][0]["message"])
        catalog = load("raise SystemExit")
        self.assertIn("SystemExit", catalog["errors"][0]["message"])

    def test_reload_replaces_functions(self):
        load()
        load("import xlfly\n@xlfly.func\ndef other(): return 1\n")
        result = calls({"kind": "call", "name": "double", "args": [1]})
        self.assertEqual(result[0]["error"]["code"], "#NAME?")

    def test_function_calls(self):
        load()
        result = calls(
            {"kind": "call", "name": "double", "args": [[[21]]]},
            {"kind": "call", "name": "TOTAL", "args": [[[1, "x"], [2, None]]]},
            {"kind": "eval", "code": "a = arg1 + 1\na * 10", "args": [4]},
            {"kind": "eval", "code": "[[x, x * x] for x in range(3)]", "args": []},
            {"kind": "eval", "code": "1/0", "args": []},
            {"kind": "eval", "code": "xlfly.XLError('#N/A')", "args": []},
        )
        self.assertEqual([r["ok"] for r in result], [True, True, True, True, False, True])
        self.assertEqual(result[0]["value"], 42)
        self.assertEqual(result[1]["value"], 3)
        self.assertEqual(result[2]["value"], 50)
        self.assertEqual(result[3]["value"], [[0, 0], [1, 1], [2, 4]])
        self.assertIn("ZeroDivisionError", result[4]["error"]["message"])
        self.assertEqual(result[5]["value"], {"$err": "#N/A"})

    def test_functions_cannot_touch_the_workbook(self):
        load()
        result = calls({"kind": "call", "name": "sneaky", "args": []})
        self.assertFalse(result[0]["ok"])
        self.assertIn("ContextError", result[0]["error"]["message"])
        self.assertEqual(self.requests, [])

    def test_functions_must_be_synchronous(self):
        load()
        result = calls({"kind": "call", "name": "not_sync", "args": []})
        self.assertIn("synchronous", result[0]["error"]["message"])

    def test_command_uses_token(self):
        load()
        result = json.loads(asyncio.run(_runtime.run_command("copy", "tok-1")))
        self.assertEqual(result, {"ok": True, "value": "B1:B2"})
        self.assertEqual([r[0] for r in self.requests], ["range_read", "range_write"])
        self.assertEqual({r[2] for r in self.requests}, {"tok-1"})
        self.assertEqual(self.requests[1][1]["values"], [[1], [2]])

    def test_unknown_command(self):
        load()
        result = json.loads(asyncio.run(_runtime.run_command("nope", "t")))
        self.assertEqual(result["error"]["code"], "NOT_FOUND")

    def test_event_dispatch(self):
        load()
        result = json.loads(asyncio.run(_runtime.dispatch_event(json.dumps({"type": "sheet_change", "address": "C3"}), "t")))
        self.assertTrue(result["ok"])
        import sys
        self.assertEqual(sys.modules["xlfly_user.main"].seen, ["C3"])


if __name__ == "__main__":
    unittest.main()
````

## 36 python/tests/test_types.py

Exercises the Python wire contract: scalar conversion, NaN, error round trips, column output, row padding, modern date conversion and scalar unwrapping. It is not a complete date-system or scientific-package test suite.

Original size: 1402 bytes. SHA-256: `b1cd6726d9ad2e6f7e742db99a34fc1ae010b9b53263fd4d66bcee23619399ae`.

````python
import datetime as dt
import math
import unittest

from xlfly._types import XLError, from_excel_date, from_wire, to_excel_date, to_wire, unwrap_single


class WireTests(unittest.TestCase):
    def test_scalars(self):
        self.assertEqual(to_wire(3), 3)
        self.assertEqual(to_wire(2.5), 2.5)
        self.assertEqual(to_wire("x"), "x")
        self.assertIs(to_wire(True), True)
        self.assertIsNone(to_wire(None))
        self.assertEqual(to_wire(math.nan), {"$err": "#NUM!"})
        self.assertEqual(to_wire(XLError("#N/A")), {"$err": "#N/A"})

    def test_sequences(self):
        self.assertEqual(to_wire([1, 2]), [[1], [2]])
        self.assertEqual(to_wire([[1], [2, 3]]), [[1, None], [2, 3]])
        self.assertEqual(to_wire([]), [[None]])

    def test_errors_round_trip(self):
        self.assertEqual(from_wire([[{"$err": "#DIV/0!"}, 1]]), [[XLError("#DIV/0!"), 1]])
        with self.assertRaises(ValueError):
            XLError("#BOGUS")

    def test_dates(self):
        self.assertEqual(to_excel_date(dt.date(2024, 1, 1)), 45292)
        self.assertEqual(to_wire(dt.datetime(2024, 1, 1, 12)), 45292.5)
        self.assertEqual(from_excel_date(45292.5), dt.datetime(2024, 1, 1, 12))

    def test_unwrap(self):
        self.assertEqual(unwrap_single([[5]]), 5)
        self.assertEqual(unwrap_single([[5, 6]]), [[5, 6]])


if __name__ == "__main__":
    unittest.main()
````

## 37 python/tests/__init__.py

An empty package marker used by unittest discovery. The full original file has zero bytes.

Original size: 0 bytes. SHA-256: `e3b0c44298fc1c149afbf4c8996fb92427ae41e4649b934ca495991b7852b855`.

````python

````

## 38 .gitignore

Keeps installed npm dependencies, generated builds, downloaded Python packages and Python bytecode out of Git. These are reconstructed from source and lockfiles.

Original size: 41 bytes. SHA-256: `f1d6d25952f4752e80da1c40e0b4ce2c917e9717b74d45b0427710e31f3ba60e`.

````text
node_modules/
dist/
vendor/
__pycache__/
````

## 39 README.md

Original project instructions retained verbatim for completeness. Its final statement about no real-Excel verification predates the newer blueprint section 10.1; use the dated blueprint and newsletter evidence table for the current snapshot status.

Original size: 4459 bytes. SHA-256: `45ebc941e1d70526bcd6caa4a4c1004541974b0027dce14ffeefbf7752874b54`.

````markdown
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
````

## 40 package-lock.json

The complete generated npm dependency lockfile. npm ci uses it for reproducible installation. It is included without abbreviation even though its repeated dependency records are not a useful line-by-line newsletter narrative.

Original size: 53554 bytes. SHA-256: `89ea09c954d362d00468802b154f821443de176d958bb9da74e4634a9a15da97`.

````json
{
  "name": "xlfly-addin",
  "version": "0.1.0",
  "lockfileVersion": 3,
  "requires": true,
  "packages": {
    "": {
      "name": "xlfly-addin",
      "version": "0.1.0",
      "license": "UNLICENSED",
      "dependencies": {
        "pyodide": "^314.0.7"
      },
      "devDependencies": {
        "@types/custom-functions-runtime": "^1.6.13",
        "@types/node": "^26.6.6",
        "@types/office-js": "^1.0.615",
        "@types/office-runtime": "^1.0.40",
        "esbuild": "^0.28.2",
        "office-addin-dev-certs": "^3.0.1",
        "office-addin-manifest": "^3.0.1",
        "typescript": "^7.0.2"
      }
    },
    "node_modules/@esbuild/aix-ppc64": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/aix-ppc64/-/aix-ppc64-0.28.2.tgz",
      "integrity": "sha512-XExcO+dvLKvVtNTibSTBej1NCAbaGhWn9Ww1ZPx80qsahhPFe/8jgWP0IchNe0F3HwkU7n8ejhH8bjonqht8mQ==",
      "cpu": [
        "ppc64"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "aix"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/android-arm": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/android-arm/-/android-arm-0.28.2.tgz",
      "integrity": "sha512-kXXoiPVVGQcnIYGOeaovwOURpniDBpSq4A03qkQ+BMQqtGG6HYap3xne9C1O1yo4TR3qxlCX5IqqmX6fFo2Lqg==",
      "cpu": [
        "arm"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "android"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/android-arm64": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/android-arm64/-/android-arm64-0.28.2.tgz",
      "integrity": "sha512-5YfKeeI8qWfBZIX+u2xZC3Zlb3Os/gLS2sbEKM+I4ZOcsWmHS2WLysCcQZDAFRslDUU5Oiq44gf6PYN1vGwG5A==",
      "cpu": [
        "arm64"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "android"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/android-x64": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/android-x64/-/android-x64-0.28.2.tgz",
      "integrity": "sha512-O387ite7SzUyCcy3JQX4P4bLtEA7bLLkx+esve5JHnyYfNTxcVpXZo9jhdB0lTKN44gztELTdU7nS8Nr16Fs1Q==",
      "cpu": [
        "x64"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "android"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/darwin-arm64": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/darwin-arm64/-/darwin-arm64-0.28.2.tgz",
      "integrity": "sha512-n4KqkOQrraxHJcgjM1RvwbigfQKIKJVpM7xp+KsxiyUSrRdIXnt73VhrPAx0fV44hgfmIVKjxMN9J1t5jySVkw==",
      "cpu": [
        "arm64"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "darwin"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/darwin-x64": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/darwin-x64/-/darwin-x64-0.28.2.tgz",
      "integrity": "sha512-uq6suIWYP37qzGddBKPw5QEQPi6HiLGsO7UmkpfyaYNQ3D+rN6w6WfwH+nuqcGXWvawGwxOEroO4YGnFh95azw==",
      "cpu": [
        "x64"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "darwin"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/freebsd-arm64": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/freebsd-arm64/-/freebsd-arm64-0.28.2.tgz",
      "integrity": "sha512-n+I0BTSRIoy+d6RPKnEVwql5UwBJolytvY4mAOIEJorKlqgPII8ix6slVVrfZ5Tnj7glIZvloylbB/EJPMWEXw==",
      "cpu": [
        "arm64"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "freebsd"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/freebsd-x64": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/freebsd-x64/-/freebsd-x64-0.28.2.tgz",
      "integrity": "sha512-78XJTJkvPs0kz2w61301PJjXl4g7q3JqiYMZ/M/yVI73EHBrCRTgkhu9oqG7vPqq+a/yadEW8aD+agKlk5xrmg==",
      "cpu": [
        "x64"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "freebsd"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/linux-arm": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/linux-arm/-/linux-arm-0.28.2.tgz",
      "integrity": "sha512-XlDnu2q5yoqems+xay6wSAcg9DDD7K9RLKZEBOMZm3ckNpJBvOX20tSfby8KfrrhINDyv9V2YVZKY/SpoGJI8w==",
      "cpu": [
        "arm"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "linux"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/linux-arm64": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/linux-arm64/-/linux-arm64-0.28.2.tgz",
      "integrity": "sha512-pW4AC0P3it8c7do9MVM4p51FzHzdM/TZrerurgRcHJ2WTa1VQ1CIq18xncfpBJw4ojkiZZrKW2yIBWBP92j6Ug==",
      "cpu": [
        "arm64"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "linux"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/linux-ia32": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/linux-ia32/-/linux-ia32-0.28.2.tgz",
      "integrity": "sha512-CYbnj78HsIeA+DhgUKgFCfvNsTHFhMMrinUrMZpDXJXKN8T3XViTZ/+wtHeVxEWY8ewSzTFN+nRmSwO2tZaLUQ==",
      "cpu": [
        "ia32"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "linux"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/linux-loong64": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/linux-loong64/-/linux-loong64-0.28.2.tgz",
      "integrity": "sha512-buwkd8nsph4R+ajRvw0qM5Hja/TXQow3ptzWO2EbG/cqcIkHloRrdlBtQlshyYGTNFvfkfJ5tpPLVkY4DtsPfQ==",
      "cpu": [
        "loong64"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "linux"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/linux-mips64el": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/linux-mips64el/-/linux-mips64el-0.28.2.tgz",
      "integrity": "sha512-ZVykbDyk7519VwiNb9Lcj9m8XM6v5V9uKPvrEMkkEedVewf+0itkhahp4HDpgERXhwLRpWFypsGbG/J8s0QjJA==",
      "cpu": [
        "mips64el"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "linux"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/linux-ppc64": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/linux-ppc64/-/linux-ppc64-0.28.2.tgz",
      "integrity": "sha512-CAXl+Dtd9UUuJd8pKKdwh6MLm3MUMiqMPmhZ3tTSXPqfyQ3vDl6R5hZdZ/kYojK4ofXtdfSv1tFq8XzWx3heNQ==",
      "cpu": [
        "ppc64"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "linux"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/linux-riscv64": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/linux-riscv64/-/linux-riscv64-0.28.2.tgz",
      "integrity": "sha512-GeXCej4IQtU1B+QlDV8W/RRvbzI3O/Stss+/bCXv4lZls5WGRtu2a+3JkA3i4qIUlMXpcHebWpF8AkJhATowuA==",
      "cpu": [
        "riscv64"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "linux"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/linux-s390x": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/linux-s390x/-/linux-s390x-0.28.2.tgz",
      "integrity": "sha512-3H1weTYZPxt/WOhByszQZybS9w5lKzUn1FDMsgEChbHWQwHYQQRfBxgCcZvPhjHfKyJjIievvMmEUawJrdY9Dg==",
      "cpu": [
        "s390x"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "linux"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/linux-x64": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/linux-x64/-/linux-x64-0.28.2.tgz",
      "integrity": "sha512-4xTZr1FUmSoQW4XIWmit3tzQrUTZM+N3P0XV8xROKYF50XfI7xeO90+1bZvNwxIufQ9hDQVRJH5YhgPVF8A/HQ==",
      "cpu": [
        "x64"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "linux"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/netbsd-arm64": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/netbsd-arm64/-/netbsd-arm64-0.28.2.tgz",
      "integrity": "sha512-sSATRjPeDBg3pdgHoQfoYBob11Kk1FGa9lui5RIHZCoCkJa9QKlvl3/vKz2usCmYYjs7ymJR/2Nnsqe+Hjt5nw==",
      "cpu": [
        "arm64"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "netbsd"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/netbsd-x64": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/netbsd-x64/-/netbsd-x64-0.28.2.tgz",
      "integrity": "sha512-lqnzCV+mM0gIADaKihiCg6ifgfU2L3h5E33rNQBN1Y4MaVGnzryzmvvf7UHxprpQdE8hpqLolJ9Rl+SkIRDpyw==",
      "cpu": [
        "x64"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "netbsd"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/openbsd-arm64": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/openbsd-arm64/-/openbsd-arm64-0.28.2.tgz",
      "integrity": "sha512-AL2qJILH7lNjrDmCQDvdxMfAUIv8KMNZOvrwAQ8i8//ntL9FflhOyMJ8OZSMBb8/AWXe3/5v5S20y3zCoZWKoQ==",
      "cpu": [
        "arm64"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "openbsd"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/openbsd-x64": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/openbsd-x64/-/openbsd-x64-0.28.2.tgz",
      "integrity": "sha512-QtiuPytchRyC4rwUKhexJdQKvDuZ6hWloi3igqPQNUJCS1/v9EiO3UTOXR6A3FoMo4fnAKbWJdqaIwhOzh8qEw==",
      "cpu": [
        "x64"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "openbsd"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/openharmony-arm64": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/openharmony-arm64/-/openharmony-arm64-0.28.2.tgz",
      "integrity": "sha512-WkhYDmpTjLvGlScA1rwjRUmhl4k8oXR3cIbtqWmELgU/dFeHHlEllxDvdWcNJV9rbzCexB5vz8gtNewWLgCT7Q==",
      "cpu": [
        "arm64"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "openharmony"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/sunos-x64": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/sunos-x64/-/sunos-x64-0.28.2.tgz",
      "integrity": "sha512-GPMSkTOtMnv2U2F8gxe4Io6qmVs+YKyp832Etqqxr0hFngmXQ3rzwytelm3GIn7T4VviRUlf3sOgBOiTdvaf7g==",
      "cpu": [
        "x64"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "sunos"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/win32-arm64": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/win32-arm64/-/win32-arm64-0.28.2.tgz",
      "integrity": "sha512-PIhhEkE9uPBleRBrQEJpUn7MBnibZzbGzYWPmY3x+YoVg/95zbjB4CxPPOQ8l5tYYM4mMaCthF8/1DIfBQQyWQ==",
      "cpu": [
        "arm64"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "win32"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/win32-ia32": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/win32-ia32/-/win32-ia32-0.28.2.tgz",
      "integrity": "sha512-YmJbfTlvU7Sdn9BB+4PRES4oB6pxgS37MAONj+hBr/cpXS1aBPKXxNnDbu+QCWPj0o9dgyxeq79g6c5P8KeuYA==",
      "cpu": [
        "ia32"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "win32"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@esbuild/win32-x64": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/@esbuild/win32-x64/-/win32-x64-0.28.2.tgz",
      "integrity": "sha512-5ebpxr3nWMzrL/rnUI755Jkuee0bHL/Gq0WTF9lvcpv73wAp5eu8MfBUgWK9bhWvZjj7yX8etf/8tI8Ney695g==",
      "cpu": [
        "x64"
      ],
      "dev": true,
      "license": "MIT",
      "optional": true,
      "os": [
        "win32"
      ],
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/@microsoft/app-manifest": {
      "version": "1.1.3",
      "resolved": "https://registry.npmjs.org/@microsoft/app-manifest/-/app-manifest-1.1.3.tgz",
      "integrity": "sha512-ensiiCOlAZFBCgj/HpYpV9jCe2Otcj21C3QZo+gYRA45zOMBGWzlDCGblnNHPqTiYesGHJr5xVtPPjCjneswpA==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "@types/fs-extra": "^11.0.1",
        "@types/strip-bom": "^4.0.1",
        "ajv": "^8.18.0",
        "ajv-draft-04": "^1.0.0",
        "ajv-formats": "^3.0.1",
        "node-fetch": "^3.3.2",
        "strip-bom": "^5.0.0"
      }
    },
    "node_modules/@microsoft/app-manifest/node_modules/node-fetch": {
      "version": "3.3.2",
      "resolved": "https://registry.npmjs.org/node-fetch/-/node-fetch-3.3.2.tgz",
      "integrity": "sha512-dRB78srN/l6gqWulah9SrxeYnxeddIG30+GOqK/9OlLVyLg3HPnr6SqOWTWOXKRwC2eGYCkZ59NNuSgvSrpgOA==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "data-uri-to-buffer": "^4.0.0",
        "fetch-blob": "^3.1.4",
        "formdata-polyfill": "^4.0.10"
      },
      "engines": {
        "node": "^12.20.0 || ^14.13.1 || >=16.0.0"
      },
      "funding": {
        "type": "opencollective",
        "url": "https://opencollective.com/node-fetch"
      }
    },
    "node_modules/@types/custom-functions-runtime": {
      "version": "1.6.13",
      "resolved": "https://registry.npmjs.org/@types/custom-functions-runtime/-/custom-functions-runtime-1.6.13.tgz",
      "integrity": "sha512-0PNbE5/sYOc5mJ1vmIDY/GFwLeuS1KZxh3eOqzCLrns+7CyoSwF0HoFpA7SeRIhQnxZKNoW7sEr03BaLHBLLbg==",
      "dev": true,
      "license": "MIT"
    },
    "node_modules/@types/emscripten": {
      "version": "1.41.6",
      "resolved": "https://registry.npmjs.org/@types/emscripten/-/emscripten-1.41.6.tgz",
      "integrity": "sha512-uN+9i8bFT5CUcZfyIEYDrSueACEyKGbUs5kC/72DGlZZoinh84sJfVV0i8UOJD1asdzkvLPBRrKs41kZ8MdEXg==",
      "license": "MIT"
    },
    "node_modules/@types/fs-extra": {
      "version": "11.0.4",
      "resolved": "https://registry.npmjs.org/@types/fs-extra/-/fs-extra-11.0.4.tgz",
      "integrity": "sha512-yTbItCNreRooED33qjunPthRcSjERP1r4MqCZc7wv0u2sUkzTFp45tgUfS5+r7FrZPdmCCNflLhVSP/o+SemsQ==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "@types/jsonfile": "*",
        "@types/node": "*"
      }
    },
    "node_modules/@types/jsonfile": {
      "version": "6.1.4",
      "resolved": "https://registry.npmjs.org/@types/jsonfile/-/jsonfile-6.1.4.tgz",
      "integrity": "sha512-D5qGUYwjvnNNextdU59/+fI+spnwtTFmyQP0h+PfIOSkNfpU6AOICUOkm4i0OnSk+NyjdPJrxCDro0sJsWlRpQ==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "@types/node": "*"
      }
    },
    "node_modules/@types/node": {
      "version": "26.6.6",
      "resolved": "https://registry.npmjs.org/@types/node/-/node-26.6.6.tgz",
      "integrity": "sha512-MzQj+8Kx3XI40OmUAE6Qoh58ZMNstNNMwn2nfd3qWsfe7nDNSqJNzes4njhz6AHwrSHOffSLwXU6yb474Q4aog==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "undici-types": "~8.9.0"
      }
    },
    "node_modules/@types/office-js": {
      "version": "1.0.615",
      "resolved": "https://registry.npmjs.org/@types/office-js/-/office-js-1.0.615.tgz",
      "integrity": "sha512-pSEiguAVkoQxSY+by3hYHo8s1Y/idmlvmpgJ1GLa5JOJrBUbtoNf5rpS3oUkaMvGrprezZ9oq9Rx3HQjLhLmbg==",
      "dev": true,
      "license": "MIT"
    },
    "node_modules/@types/office-runtime": {
      "version": "1.0.40",
      "resolved": "https://registry.npmjs.org/@types/office-runtime/-/office-runtime-1.0.40.tgz",
      "integrity": "sha512-eFEhqFnYP5xmkZKtVz0Mk3AsD/sLADkhQqGbG1IWylfC2rj+5ZjOZEmD7k5j77vTcMYN0hF1cQy9OW984HCh0g==",
      "dev": true,
      "license": "MIT"
    },
    "node_modules/@types/strip-bom": {
      "version": "4.0.1",
      "resolved": "https://registry.npmjs.org/@types/strip-bom/-/strip-bom-4.0.1.tgz",
      "integrity": "sha512-d3RZcMmYRuL5f+jG5YM5ISs9z/HCwI2xjqXxhLon54dlpBWfVLVStOFTalO7UcX7RXaIJ6oylqOTg3Cbu6zU1Q==",
      "deprecated": "This is a stub types definition. strip-bom provides its own type definitions, so you do not need this installed.",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "strip-bom": "*"
      }
    },
    "node_modules/@typescript/typescript-aix-ppc64": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/@typescript/typescript-aix-ppc64/-/typescript-aix-ppc64-7.0.2.tgz",
      "integrity": "sha512-MTKKkWB7p/0E9xi1d1tHtZ5PiLkGEMIq88pK2CubZjOsLtYTLqhgIgi6zepFa+9GHZ6h05NMCkQxGKiPXMxXtQ==",
      "cpu": [
        "ppc64"
      ],
      "dev": true,
      "license": "Apache-2.0",
      "optional": true,
      "os": [
        "aix"
      ],
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/@typescript/typescript-darwin-arm64": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/@typescript/typescript-darwin-arm64/-/typescript-darwin-arm64-7.0.2.tgz",
      "integrity": "sha512-gowzar9MwS/aRWp6f3a4KUqzRjAZjOsmGNCM6LcTgXum+dBfgsBVMN+AgvOCCbguXyick6LJhpBszxMebJ8syA==",
      "cpu": [
        "arm64"
      ],
      "dev": true,
      "license": "Apache-2.0",
      "optional": true,
      "os": [
        "darwin"
      ],
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/@typescript/typescript-darwin-x64": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/@typescript/typescript-darwin-x64/-/typescript-darwin-x64-7.0.2.tgz",
      "integrity": "sha512-SZ9xZInqApNlNGc9s0W1VSsktYSOe9cFqNOIqmN1Gs8SmkjKZYFt017G4VwPxASInODuAdbTW7sXiFUf893RgA==",
      "cpu": [
        "x64"
      ],
      "dev": true,
      "license": "Apache-2.0",
      "optional": true,
      "os": [
        "darwin"
      ],
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/@typescript/typescript-freebsd-arm64": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/@typescript/typescript-freebsd-arm64/-/typescript-freebsd-arm64-7.0.2.tgz",
      "integrity": "sha512-W5NH4y/J0plIIS5b2xvTEkU7JFxyqdMAOgf+Ilhl0vHQXKO5dZoxd+C/jEtq56c4F3wk71RB4BMRQ2XdI+bwYQ==",
      "cpu": [
        "arm64"
      ],
      "dev": true,
      "license": "Apache-2.0",
      "optional": true,
      "os": [
        "freebsd"
      ],
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/@typescript/typescript-freebsd-x64": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/@typescript/typescript-freebsd-x64/-/typescript-freebsd-x64-7.0.2.tgz",
      "integrity": "sha512-UMGDx5sTpzNw3WiPebH7l90IWfJggEd+egHt/q6p7/Cm3zqoV7VxkGXt+3DxPIw8CcmvAB0j3sVVfbhX+M4Tpw==",
      "cpu": [
        "x64"
      ],
      "dev": true,
      "license": "Apache-2.0",
      "optional": true,
      "os": [
        "freebsd"
      ],
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/@typescript/typescript-linux-arm": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/@typescript/typescript-linux-arm/-/typescript-linux-arm-7.0.2.tgz",
      "integrity": "sha512-gffT3xPz9sR7j/YJExkyPntrI0P2EP9XbOyWzth2/Gs0RstK+90RBcO0ncXoXy/beYll1SXw846Nf2zdnEz0QQ==",
      "cpu": [
        "arm"
      ],
      "dev": true,
      "license": "Apache-2.0",
      "optional": true,
      "os": [
        "linux"
      ],
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/@typescript/typescript-linux-arm64": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/@typescript/typescript-linux-arm64/-/typescript-linux-arm64-7.0.2.tgz",
      "integrity": "sha512-Qh4eU4/y3yDjnfjjyPYihMj5/ODIlmt+Bzu17OI+fiSRDW57QmU5SiN63exPRNJPKUzcc1INa1NXdrJ+MqHjUQ==",
      "cpu": [
        "arm64"
      ],
      "dev": true,
      "license": "Apache-2.0",
      "optional": true,
      "os": [
        "linux"
      ],
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/@typescript/typescript-linux-loong64": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/@typescript/typescript-linux-loong64/-/typescript-linux-loong64-7.0.2.tgz",
      "integrity": "sha512-uEHck9i8hoAzXPiYRib1O7miOnz23SxIeVl6F4LXox+qov1K35jHcEW6VHKvZI+pyvl7fZEP4MCU5LYvIq1GuQ==",
      "cpu": [
        "loong64"
      ],
      "dev": true,
      "license": "Apache-2.0",
      "optional": true,
      "os": [
        "linux"
      ],
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/@typescript/typescript-linux-mips64el": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/@typescript/typescript-linux-mips64el/-/typescript-linux-mips64el-7.0.2.tgz",
      "integrity": "sha512-R4KvAMnE43W5Qeqb0Ly56O3mWMWIAgsMyz36DCaycd5nbg/9kzm0liw3JocfRqyJY0KPmzFjbswozXyW0DnIYA==",
      "cpu": [
        "mips64el"
      ],
      "dev": true,
      "license": "Apache-2.0",
      "optional": true,
      "os": [
        "linux"
      ],
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/@typescript/typescript-linux-ppc64": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/@typescript/typescript-linux-ppc64/-/typescript-linux-ppc64-7.0.2.tgz",
      "integrity": "sha512-DORx5b3sd/4S7eayxm4FQv+A7CrkUIGRaHiwI8oiHTAI1fAPWhF4J0vAlkC8biAlHSVVwxMQ3tjZ2/DVbnQiiA==",
      "cpu": [
        "ppc64"
      ],
      "dev": true,
      "license": "Apache-2.0",
      "optional": true,
      "os": [
        "linux"
      ],
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/@typescript/typescript-linux-riscv64": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/@typescript/typescript-linux-riscv64/-/typescript-linux-riscv64-7.0.2.tgz",
      "integrity": "sha512-wf0jqEDOjrPRnKwYRyyJDRo11KMbvMFrU+q4zqKyChODBzvlkbhNQfKvLxQCcwTpdDaXSHZTVuh0JoCrKCUMHQ==",
      "cpu": [
        "riscv64"
      ],
      "dev": true,
      "license": "Apache-2.0",
      "optional": true,
      "os": [
        "linux"
      ],
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/@typescript/typescript-linux-s390x": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/@typescript/typescript-linux-s390x/-/typescript-linux-s390x-7.0.2.tgz",
      "integrity": "sha512-IkwJc3L7yhytWd/ewjyxNDfOmswCm9GWMJT/ue/dU4aZNbwZeYAetq42VyLmsmSjvoX7z74X6ZaYCtzAr0EuGw==",
      "cpu": [
        "s390x"
      ],
      "dev": true,
      "license": "Apache-2.0",
      "optional": true,
      "os": [
        "linux"
      ],
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/@typescript/typescript-linux-x64": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/@typescript/typescript-linux-x64/-/typescript-linux-x64-7.0.2.tgz",
      "integrity": "sha512-EYdf2cNg7rgCWJnxCdJ+F3V39O8ihb37eHAu1LK8oAFizgTQbPOK7zHHXbPt8rX24COqODXeI3sIf0fCXG7H/A==",
      "cpu": [
        "x64"
      ],
      "dev": true,
      "license": "Apache-2.0",
      "optional": true,
      "os": [
        "linux"
      ],
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/@typescript/typescript-netbsd-arm64": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/@typescript/typescript-netbsd-arm64/-/typescript-netbsd-arm64-7.0.2.tgz",
      "integrity": "sha512-+polYF4MF04aPpO5FTkHran9yUQDSXqy5GiSDKpsll5jy3l3+g9QLhpf39T+ePtefhXLOGrLl0QIjkQP6VnelA==",
      "cpu": [
        "arm64"
      ],
      "dev": true,
      "license": "Apache-2.0",
      "optional": true,
      "os": [
        "netbsd"
      ],
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/@typescript/typescript-netbsd-x64": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/@typescript/typescript-netbsd-x64/-/typescript-netbsd-x64-7.0.2.tgz",
      "integrity": "sha512-8YIT0EHM/3dq10ZOVF/A7pc/YSMtbcecct4rWtexrnSCHOPcpC2KTLXfTCR6vDpnSiY12heNb1GiN/wu+T/FyA==",
      "cpu": [
        "x64"
      ],
      "dev": true,
      "license": "Apache-2.0",
      "optional": true,
      "os": [
        "netbsd"
      ],
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/@typescript/typescript-openbsd-arm64": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/@typescript/typescript-openbsd-arm64/-/typescript-openbsd-arm64-7.0.2.tgz",
      "integrity": "sha512-APT8+ClYnuYm1u9+kgGXoMj2VzWzcymwh2gNSQVySHfkRDGOTVkoWLjCmOQSaO+PoqQ57B0flRp9SA+7GnnkzQ==",
      "cpu": [
        "arm64"
      ],
      "dev": true,
      "license": "Apache-2.0",
      "optional": true,
      "os": [
        "openbsd"
      ],
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/@typescript/typescript-openbsd-x64": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/@typescript/typescript-openbsd-x64/-/typescript-openbsd-x64-7.0.2.tgz",
      "integrity": "sha512-yX7s+Q0Dln0Dt9tEzZsAjXXR/+ytBM7AlglaqyeMPxQszJ1JhlJdZ6jLA+IzldHtflX81em7lDao1xXu+aRRkg==",
      "cpu": [
        "x64"
      ],
      "dev": true,
      "license": "Apache-2.0",
      "optional": true,
      "os": [
        "openbsd"
      ],
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/@typescript/typescript-sunos-x64": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/@typescript/typescript-sunos-x64/-/typescript-sunos-x64-7.0.2.tgz",
      "integrity": "sha512-dLJDGaLZ1D4HPQn62u1n8mBDkJREwMsAkCdkwd4Ieqw+x3TUyTsqY0YiBCtE6H6OzzgGk3iuZ3vFWRS+E8/d1g==",
      "cpu": [
        "x64"
      ],
      "dev": true,
      "license": "Apache-2.0",
      "optional": true,
      "os": [
        "sunos"
      ],
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/@typescript/typescript-win32-arm64": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/@typescript/typescript-win32-arm64/-/typescript-win32-arm64-7.0.2.tgz",
      "integrity": "sha512-Gyl1Vy6OsWesLzmq+EP0Fb7b4Nid5232AvcA2SFcdYreldpNtYFFofPjnt62y9hQy7VTaZp65ICJjuAQRaVcIQ==",
      "cpu": [
        "arm64"
      ],
      "dev": true,
      "license": "Apache-2.0",
      "optional": true,
      "os": [
        "win32"
      ],
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/@typescript/typescript-win32-x64": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/@typescript/typescript-win32-x64/-/typescript-win32-x64-7.0.2.tgz",
      "integrity": "sha512-0BQ3HkAHHlKLSp1qRvf3SUhGpGsDuhB/jgFw75guyqbxJqEaS0Cw/VFO8i2nHglJUzQCRtMMR/IBAKE3ETMC4g==",
      "cpu": [
        "x64"
      ],
      "dev": true,
      "license": "Apache-2.0",
      "optional": true,
      "os": [
        "win32"
      ],
      "engines": {
        "node": ">=16.20.0"
      }
    },
    "node_modules/adm-zip": {
      "version": "0.6.1",
      "resolved": "https://registry.npmjs.org/adm-zip/-/adm-zip-0.6.1.tgz",
      "integrity": "sha512-Xwrja8nx9e5o2N1my4DsKCeKpdrnACyr1wtbPxBDgGzKzKyE9kRtBFA8mWldI+RVlD7CBZNWY/wQ2+ydwOR6kQ==",
      "dev": true,
      "license": "MIT",
      "engines": {
        "node": ">=14.0"
      }
    },
    "node_modules/ajv": {
      "version": "8.20.0",
      "resolved": "https://registry.npmjs.org/ajv/-/ajv-8.20.0.tgz",
      "integrity": "sha512-Thbli+OlOj+iMPYFBVBfJ3OmCAnaSyNn4M1vz9T6Gka5Jt9ba/HIR56joy65tY6kx/FCF5VXNB819Y7/GUrBGA==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "fast-deep-equal": "^3.1.3",
        "fast-uri": "^3.0.1",
        "json-schema-traverse": "^1.0.0",
        "require-from-string": "^2.0.2"
      },
      "funding": {
        "type": "github",
        "url": "https://github.com/sponsors/epoberezkin"
      }
    },
    "node_modules/ajv-draft-04": {
      "version": "1.0.0",
      "resolved": "https://registry.npmjs.org/ajv-draft-04/-/ajv-draft-04-1.0.0.tgz",
      "integrity": "sha512-mv00Te6nmYbRp5DCwclxtt7yV/joXJPGS7nM+97GdxvuttCOfgI3K4U25zboyeX0O+myI8ERluxQe5wljMmVIw==",
      "dev": true,
      "license": "MIT",
      "peerDependencies": {
        "ajv": "^8.5.0"
      },
      "peerDependenciesMeta": {
        "ajv": {
          "optional": true
        }
      }
    },
    "node_modules/ajv-formats": {
      "version": "3.0.1",
      "resolved": "https://registry.npmjs.org/ajv-formats/-/ajv-formats-3.0.1.tgz",
      "integrity": "sha512-8iUql50EUR+uUcdRQ3HDqa6EVyo3docL8g5WJ3FNcWmu62IbkGUue/pEyLBW8VGKKucTPgqeks4fIU1DA4yowQ==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "ajv": "^8.0.0"
      },
      "peerDependencies": {
        "ajv": "^8.0.0"
      },
      "peerDependenciesMeta": {
        "ajv": {
          "optional": true
        }
      }
    },
    "node_modules/ansi-styles": {
      "version": "3.2.1",
      "resolved": "https://registry.npmjs.org/ansi-styles/-/ansi-styles-3.2.1.tgz",
      "integrity": "sha512-VT0ZI6kZRdTh8YyJw3SMbYm/u+NqfsAxEpWO0Pf9sq8/e94WxxOpPKx9FR1FlyCtOVDNOQ+8ntlqFxiRc+r5qA==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "color-convert": "^1.9.0"
      },
      "engines": {
        "node": ">=4"
      }
    },
    "node_modules/chalk": {
      "version": "2.4.2",
      "resolved": "https://registry.npmjs.org/chalk/-/chalk-2.4.2.tgz",
      "integrity": "sha512-Mti+f9lpJNcwF4tWV8/OrTTtF1gZi+f8FqlyAdouralcFWFQWF2+NgCHShjkCb+IFBLq9buZwE1xckQU4peSuQ==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "ansi-styles": "^3.2.1",
        "escape-string-regexp": "^1.0.5",
        "supports-color": "^5.3.0"
      },
      "engines": {
        "node": ">=4"
      }
    },
    "node_modules/color-convert": {
      "version": "1.9.3",
      "resolved": "https://registry.npmjs.org/color-convert/-/color-convert-1.9.3.tgz",
      "integrity": "sha512-QfAUtd+vFdAtFQcC8CCyYt1fYWxSqAiK2cSD6zDB8N3cpsEBAvRxp9zOGg6G/SHHJYAT88/az/IuDGALsNVbGg==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "color-name": "1.1.3"
      }
    },
    "node_modules/color-name": {
      "version": "1.1.3",
      "resolved": "https://registry.npmjs.org/color-name/-/color-name-1.1.3.tgz",
      "integrity": "sha512-72fSenhMw2HZMTVHeCA9KCmpEIbzWiQsjN+BHcBbS9vr1mtt+vJjPdksIBNUmKAW8TFUDPJK5SUU3QhE9NEXDw==",
      "dev": true,
      "license": "MIT"
    },
    "node_modules/commander": {
      "version": "13.1.0",
      "resolved": "https://registry.npmjs.org/commander/-/commander-13.1.0.tgz",
      "integrity": "sha512-/rFeCpNJQbhSZjGVwO9RFV3xPqbnERS8MmIQzCtD/zl6gpJuV/bMLuN92oG3F7d8oDEHHRrujSXNUr8fpjntKw==",
      "dev": true,
      "license": "MIT",
      "engines": {
        "node": ">=18"
      }
    },
    "node_modules/data-uri-to-buffer": {
      "version": "4.0.1",
      "resolved": "https://registry.npmjs.org/data-uri-to-buffer/-/data-uri-to-buffer-4.0.1.tgz",
      "integrity": "sha512-0R9ikRb668HB7QDxT1vkpuUBtqc53YyAwMwGeUFKRojY/NWKvdZ+9UYtRfGmhqNbRkTSVpMbmyhXipFFv2cb/A==",
      "dev": true,
      "license": "MIT",
      "engines": {
        "node": ">= 12"
      }
    },
    "node_modules/esbuild": {
      "version": "0.28.2",
      "resolved": "https://registry.npmjs.org/esbuild/-/esbuild-0.28.2.tgz",
      "integrity": "sha512-HKVLS8dvII+xoKW9kmqxbRKrnWEXfJJr/FZhhJmiqIB0e053QNYFqOBouTMO/k5sID4MvCiUCvv8b9M4h32wIA==",
      "dev": true,
      "hasInstallScript": true,
      "license": "MIT",
      "bin": {
        "esbuild": "bin/esbuild"
      },
      "engines": {
        "node": ">=18"
      },
      "optionalDependencies": {
        "@esbuild/aix-ppc64": "0.28.2",
        "@esbuild/android-arm": "0.28.2",
        "@esbuild/android-arm64": "0.28.2",
        "@esbuild/android-x64": "0.28.2",
        "@esbuild/darwin-arm64": "0.28.2",
        "@esbuild/darwin-x64": "0.28.2",
        "@esbuild/freebsd-arm64": "0.28.2",
        "@esbuild/freebsd-x64": "0.28.2",
        "@esbuild/linux-arm": "0.28.2",
        "@esbuild/linux-arm64": "0.28.2",
        "@esbuild/linux-ia32": "0.28.2",
        "@esbuild/linux-loong64": "0.28.2",
        "@esbuild/linux-mips64el": "0.28.2",
        "@esbuild/linux-ppc64": "0.28.2",
        "@esbuild/linux-riscv64": "0.28.2",
        "@esbuild/linux-s390x": "0.28.2",
        "@esbuild/linux-x64": "0.28.2",
        "@esbuild/netbsd-arm64": "0.28.2",
        "@esbuild/netbsd-x64": "0.28.2",
        "@esbuild/openbsd-arm64": "0.28.2",
        "@esbuild/openbsd-x64": "0.28.2",
        "@esbuild/openharmony-arm64": "0.28.2",
        "@esbuild/sunos-x64": "0.28.2",
        "@esbuild/win32-arm64": "0.28.2",
        "@esbuild/win32-ia32": "0.28.2",
        "@esbuild/win32-x64": "0.28.2"
      }
    },
    "node_modules/escape-string-regexp": {
      "version": "1.0.5",
      "resolved": "https://registry.npmjs.org/escape-string-regexp/-/escape-string-regexp-1.0.5.tgz",
      "integrity": "sha512-vbRorB5FUQWvla16U8R/qgaFIya2qGzwDrNmCZuYKrbdSUMG6I1ZCGQRefkRVhuOkIGVne7BQ35DSfo1qvJqFg==",
      "dev": true,
      "license": "MIT",
      "engines": {
        "node": ">=0.8.0"
      }
    },
    "node_modules/fast-deep-equal": {
      "version": "3.1.3",
      "resolved": "https://registry.npmjs.org/fast-deep-equal/-/fast-deep-equal-3.1.3.tgz",
      "integrity": "sha512-f3qQ9oQy9j2AhBe/H9VC91wLmKBCCU/gDOnKNAYG5hswO7BLKj09Hc5HYNz9cGI++xlpDCIgDaitVs03ATR84Q==",
      "dev": true,
      "license": "MIT"
    },
    "node_modules/fast-uri": {
      "version": "3.1.9",
      "resolved": "https://registry.npmjs.org/fast-uri/-/fast-uri-3.1.9.tgz",
      "integrity": "sha512-hS/0Wt3Ya043oKXs0sIxrMpKQITXhMLIrVAKKXi4gozGC9v7YRWVnXIZzjNPTO02M9BdxLCWqK1yRXgNziVyJQ==",
      "dev": true,
      "funding": [
        {
          "type": "github",
          "url": "https://github.com/sponsors/fastify"
        },
        {
          "type": "opencollective",
          "url": "https://opencollective.com/fastify"
        }
      ],
      "license": "BSD-3-Clause"
    },
    "node_modules/fetch-blob": {
      "version": "3.2.0",
      "resolved": "https://registry.npmjs.org/fetch-blob/-/fetch-blob-3.2.0.tgz",
      "integrity": "sha512-7yAQpD2UMJzLi1Dqv7qFYnPbaPx7ZfFK6PiIxQ4PfkGPyNyl2Ugx+a/umUonmKqjhM4DnfbMvdX6otXq83soQQ==",
      "dev": true,
      "funding": [
        {
          "type": "github",
          "url": "https://github.com/sponsors/jimmywarting"
        },
        {
          "type": "paypal",
          "url": "https://paypal.me/jimmywarting"
        }
      ],
      "license": "MIT",
      "dependencies": {
        "node-domexception": "^1.0.0",
        "web-streams-polyfill": "^3.0.3"
      },
      "engines": {
        "node": "^12.20 || >= 14.13"
      }
    },
    "node_modules/formdata-polyfill": {
      "version": "4.0.10",
      "resolved": "https://registry.npmjs.org/formdata-polyfill/-/formdata-polyfill-4.0.10.tgz",
      "integrity": "sha512-buewHzMvYL29jdeQTVILecSaZKnt/RJWjoZCF5OW60Z67/GmSLBkOFM7qh1PI3zFNtJbaZL5eQu1vLfazOwj4g==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "fetch-blob": "^3.1.2"
      },
      "engines": {
        "node": ">=12.20.0"
      }
    },
    "node_modules/fs-extra": {
      "version": "11.4.1",
      "resolved": "https://registry.npmjs.org/fs-extra/-/fs-extra-11.4.1.tgz",
      "integrity": "sha512-KYAb4c9BJQI6QqGKthV68OHe0badztdXJWKo0WtBA9IuCFPTKvE5ZdUBglP833aMjhaSPNO4A5j/EkzZtGlKjA==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "graceful-fs": "^4.2.0",
        "jsonfile": "^6.0.1",
        "universalify": "^2.0.0"
      },
      "engines": {
        "node": ">=14.14"
      }
    },
    "node_modules/graceful-fs": {
      "version": "4.2.11",
      "resolved": "https://registry.npmjs.org/graceful-fs/-/graceful-fs-4.2.11.tgz",
      "integrity": "sha512-RbJ5/jmFcNNCcDV5o9eTnBLJ/HszWV0P73bc+Ff4nS/rJj+YaS6IGyiOL0VoBYX+l1Wrl3k63h/KrH+nhJ0XvQ==",
      "dev": true,
      "license": "ISC"
    },
    "node_modules/has-flag": {
      "version": "3.0.0",
      "resolved": "https://registry.npmjs.org/has-flag/-/has-flag-3.0.0.tgz",
      "integrity": "sha512-sKJf1+ceQBr4SMkvQnBDNDtf4TXpVhVGateu0t918bl30FnbE2m4vNLX+VWe/dpjlb+HugGYzW7uQXH98HPEYw==",
      "dev": true,
      "license": "MIT",
      "engines": {
        "node": ">=4"
      }
    },
    "node_modules/json-parse-even-better-errors": {
      "version": "2.3.1",
      "resolved": "https://registry.npmjs.org/json-parse-even-better-errors/-/json-parse-even-better-errors-2.3.1.tgz",
      "integrity": "sha512-xyFwyhro/JEof6Ghe2iz2NcXoj2sloNsWr/XsERDK/oiPCfaNhl5ONfp+jQdAZRQQ0IJWNzH9zIZF7li91kh2w==",
      "dev": true,
      "license": "MIT"
    },
    "node_modules/json-schema-traverse": {
      "version": "1.0.0",
      "resolved": "https://registry.npmjs.org/json-schema-traverse/-/json-schema-traverse-1.0.0.tgz",
      "integrity": "sha512-NM8/P9n3XjXhIZn1lLhkFaACTOURQXjWhV4BA/RnOv8xvgqtqpAX9IO4mRQxSx1Rlo4tqzeqb0sOlruaOy3dug==",
      "dev": true,
      "license": "MIT"
    },
    "node_modules/jsonfile": {
      "version": "6.2.1",
      "resolved": "https://registry.npmjs.org/jsonfile/-/jsonfile-6.2.1.tgz",
      "integrity": "sha512-zwOTdL3rFQ/lRdBnntKVOX6k5cKJwEc1HdilT71BWEu7J41gXIB2MRp+vxduPSwZJPWBxEzv4yH1wYLJGUHX4Q==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "universalify": "^2.0.0"
      },
      "optionalDependencies": {
        "graceful-fs": "^4.1.6"
      }
    },
    "node_modules/mkcert": {
      "version": "3.2.0",
      "resolved": "https://registry.npmjs.org/mkcert/-/mkcert-3.2.0.tgz",
      "integrity": "sha512-026Eivq9RoOjOuLJGzbhGwXUAjBxRX11Z7Jbm4/7lqT/Av+XNy9SPrJte6+UpEt7i+W3e/HZYxQqlQcqXZWSzg==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "commander": "^11.0.0",
        "node-forge": "^1.3.1"
      },
      "bin": {
        "mkcert": "dist/cli.js"
      },
      "engines": {
        "node": ">=16"
      }
    },
    "node_modules/mkcert/node_modules/commander": {
      "version": "11.1.0",
      "resolved": "https://registry.npmjs.org/commander/-/commander-11.1.0.tgz",
      "integrity": "sha512-yPVavfyCcRhmorC7rWlkHn15b4wDVgVmBA7kV4QVBsF7kv/9TKJAbAXVTxvTnwP8HHKjRCJDClKbciiYS7p0DQ==",
      "dev": true,
      "license": "MIT",
      "engines": {
        "node": ">=16"
      }
    },
    "node_modules/node-domexception": {
      "version": "1.0.0",
      "resolved": "https://registry.npmjs.org/node-domexception/-/node-domexception-1.0.0.tgz",
      "integrity": "sha512-/jKZoMpw0F8GRwl4/eLROPA3cfcXtLApP0QzLmUT/HuPCZWyB7IY9ZrMeKw2O/nFIqPQB3PVM9aYm0F312AXDQ==",
      "deprecated": "Use your platform's native DOMException instead",
      "dev": true,
      "funding": [
        {
          "type": "github",
          "url": "https://github.com/sponsors/jimmywarting"
        },
        {
          "type": "github",
          "url": "https://paypal.me/jimmywarting"
        }
      ],
      "license": "MIT",
      "engines": {
        "node": ">=10.5.0"
      }
    },
    "node_modules/node-fetch": {
      "version": "2.7.0",
      "resolved": "https://registry.npmjs.org/node-fetch/-/node-fetch-2.7.0.tgz",
      "integrity": "sha512-c4FRfUm/dbcWZ7U+1Wq0AwCyFL+3nt2bEw05wfxSz+DWpWsitgmSgYmy2dQdWyKC1694ELPqMs/YzUSNozLt8A==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "whatwg-url": "^5.0.0"
      },
      "engines": {
        "node": "4.x || >=6.0.0"
      },
      "peerDependencies": {
        "encoding": "^0.1.0"
      },
      "peerDependenciesMeta": {
        "encoding": {
          "optional": true
        }
      }
    },
    "node_modules/node-forge": {
      "version": "1.4.2",
      "resolved": "https://registry.npmjs.org/node-forge/-/node-forge-1.4.2.tgz",
      "integrity": "sha512-rRtAR5XLJ8gCKd9du4qrKAsgCG1ObGP4gLeNJMa4vYSeq4bmBMOJk5IHkrviBV0q7uoAJlVd8HEQYQSVUb4FGQ==",
      "dev": true,
      "license": "(BSD-3-Clause OR GPL-2.0)",
      "engines": {
        "node": ">= 6.13.0"
      }
    },
    "node_modules/npm-normalize-package-bin": {
      "version": "1.0.1",
      "resolved": "https://registry.npmjs.org/npm-normalize-package-bin/-/npm-normalize-package-bin-1.0.1.tgz",
      "integrity": "sha512-EPfafl6JL5/rU+ot6P3gRSCpPDW5VmIzX959Ob1+ySFUuuYHWHekXpwdUZcKP5C+DS4GEtdJluwBjnsNDl+fSA==",
      "dev": true,
      "license": "ISC"
    },
    "node_modules/office-addin-cli": {
      "version": "3.0.1",
      "resolved": "https://registry.npmjs.org/office-addin-cli/-/office-addin-cli-3.0.1.tgz",
      "integrity": "sha512-8bklQAMMf+ZLUr+yfNTNNA73feizf0qOapXbToZ0hr4oiA9rTEebIbO29Uh80gqQNiPRaGGny7IZq/M2RG60ug==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "commander": "^13.0.0",
        "node-fetch": "^2.6.1",
        "read-package-json-fast": "^2.0.2"
      },
      "bin": {
        "office-addin-cli": "cli.js"
      }
    },
    "node_modules/office-addin-dev-certs": {
      "version": "3.0.1",
      "resolved": "https://registry.npmjs.org/office-addin-dev-certs/-/office-addin-dev-certs-3.0.1.tgz",
      "integrity": "sha512-4ndnqbeQLUexNvwpx2HRIQAsh5hAvv2phgh/YQcZYUrr5gXYCtD//pK35ZVSxCb/L7CNI2ExyaLJ6OmBk5y/Nw==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "commander": "^13.0.0",
        "fs-extra": "^11.2.0",
        "mkcert": "^3.2.0",
        "office-addin-cli": "^3.0.1",
        "office-addin-usage-data": "^3.0.1"
      },
      "bin": {
        "office-addin-dev-certs": "cli.js"
      }
    },
    "node_modules/office-addin-manifest": {
      "version": "3.0.1",
      "resolved": "https://registry.npmjs.org/office-addin-manifest/-/office-addin-manifest-3.0.1.tgz",
      "integrity": "sha512-1gHBxW+Bri23UCNK7x3D1qhogsSUFiRE5q8NR+lI1MwKE35Jq8nUrNKO0ZgT7WiEfxl4ZcZlmylGrsvMVoooLw==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "@microsoft/app-manifest": "^1.0.2",
        "adm-zip": "^0.6.1",
        "chalk": "^2.4.2",
        "commander": "^13.0.0",
        "fs-extra": "^7.0.1",
        "node-fetch": "^2.6.1",
        "office-addin-usage-data": "^3.0.1",
        "uuid": "^14.0.1",
        "xml2js": "^0.5.0"
      },
      "bin": {
        "office-addin-manifest": "cli.js"
      }
    },
    "node_modules/office-addin-manifest/node_modules/fs-extra": {
      "version": "7.0.1",
      "resolved": "https://registry.npmjs.org/fs-extra/-/fs-extra-7.0.1.tgz",
      "integrity": "sha512-YJDaCJZEnBmcbw13fvdAM9AwNOJwOzrE4pqMqBq5nFiEqXUqHwlK4B+3pUw6JNvfSPtX05xFHtYy/1ni01eGCw==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "graceful-fs": "^4.1.2",
        "jsonfile": "^4.0.0",
        "universalify": "^0.1.0"
      },
      "engines": {
        "node": ">=6 <7 || >=8"
      }
    },
    "node_modules/office-addin-manifest/node_modules/jsonfile": {
      "version": "4.0.0",
      "resolved": "https://registry.npmjs.org/jsonfile/-/jsonfile-4.0.0.tgz",
      "integrity": "sha512-m6F1R3z8jjlf2imQHS2Qez5sjKWQzbuuhuJ/FKYFRZvPE3PuHcSMVZzfsLhGVOkfd20obL5SWEBew5ShlquNxg==",
      "dev": true,
      "license": "MIT",
      "optionalDependencies": {
        "graceful-fs": "^4.1.6"
      }
    },
    "node_modules/office-addin-manifest/node_modules/universalify": {
      "version": "0.1.2",
      "resolved": "https://registry.npmjs.org/universalify/-/universalify-0.1.2.tgz",
      "integrity": "sha512-rBJeI5CXAlmy1pV+617WB9J63U6XcazHHF2f2dbJix4XzpUF0RS3Zbj0FGIOCAva5P/d/GBOYaACQ1w+0azUkg==",
      "dev": true,
      "license": "MIT",
      "engines": {
        "node": ">= 4.0.0"
      }
    },
    "node_modules/office-addin-usage-data": {
      "version": "3.0.1",
      "resolved": "https://registry.npmjs.org/office-addin-usage-data/-/office-addin-usage-data-3.0.1.tgz",
      "integrity": "sha512-nso5kikrl9MH+4aqDiIZtc5i8RCd6VsoQZu4CrDm75wLLXYoBnwl3Wux+XvM42AytJi94qyFuR5rp/+Er5o6Sg==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "commander": "^13.0.0"
      },
      "bin": {
        "office-addin-usage-data": "cli.js"
      }
    },
    "node_modules/pyodide": {
      "version": "314.0.7",
      "resolved": "https://registry.npmjs.org/pyodide/-/pyodide-314.0.7.tgz",
      "integrity": "sha512-0YvXxEhfEdpLfb/XkM2BFAeMROq0iMUX2bzzH9pOttyMcWkwq+HbE5uyuGD82LN7y2q+SNvi/6V5JEsOlD2R1A==",
      "license": "MPL-2.0",
      "dependencies": {
        "@types/emscripten": "^1.41.4",
        "ws": "^8.5.0"
      },
      "engines": {
        "node": ">=18.0.0"
      }
    },
    "node_modules/read-package-json-fast": {
      "version": "2.0.3",
      "resolved": "https://registry.npmjs.org/read-package-json-fast/-/read-package-json-fast-2.0.3.tgz",
      "integrity": "sha512-W/BKtbL+dUjTuRL2vziuYhp76s5HZ9qQhd/dKfWIZveD0O40453QNyZhC0e63lqZrAQ4jiOapVoeJ7JrszenQQ==",
      "dev": true,
      "license": "ISC",
      "dependencies": {
        "json-parse-even-better-errors": "^2.3.0",
        "npm-normalize-package-bin": "^1.0.1"
      },
      "engines": {
        "node": ">=10"
      }
    },
    "node_modules/require-from-string": {
      "version": "2.0.2",
      "resolved": "https://registry.npmjs.org/require-from-string/-/require-from-string-2.0.2.tgz",
      "integrity": "sha512-Xf0nWe6RseziFMu+Ap9biiUbmplq6S9/p+7w7YXP/JBHhrUDDUhwa+vANyubuqfZWTveU//DYVGsDG7RKL/vEw==",
      "dev": true,
      "license": "MIT",
      "engines": {
        "node": ">=0.10.0"
      }
    },
    "node_modules/sax": {
      "version": "1.6.1",
      "resolved": "https://registry.npmjs.org/sax/-/sax-1.6.1.tgz",
      "integrity": "sha512-42tBVwLWnaQvW5zc4HbZrTuWccECCZfBi92FDuwtqxasH+JbPB3/FOKb1m222K42R4WxuxzzMsTswfzgtSu64Q==",
      "dev": true,
      "license": "BlueOak-1.0.0",
      "engines": {
        "node": ">=11.0.0"
      }
    },
    "node_modules/strip-bom": {
      "version": "5.0.0",
      "resolved": "https://registry.npmjs.org/strip-bom/-/strip-bom-5.0.0.tgz",
      "integrity": "sha512-p+byADHF7SzEcVnLvc/r3uognM1hUhObuHXxJcgLCfD194XAkaLbjq3Wzb0N5G2tgIjH0dgT708Z51QxMeu60A==",
      "dev": true,
      "license": "MIT",
      "engines": {
        "node": ">=12"
      },
      "funding": {
        "url": "https://github.com/sponsors/sindresorhus"
      }
    },
    "node_modules/supports-color": {
      "version": "5.5.0",
      "resolved": "https://registry.npmjs.org/supports-color/-/supports-color-5.5.0.tgz",
      "integrity": "sha512-QjVjwdXIt408MIiAqCX4oUKsgU2EqAGzs2Ppkm4aQYbjm+ZEWEcW4SfFNTr4uMNZma0ey4f5lgLrkB0aX0QMow==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "has-flag": "^3.0.0"
      },
      "engines": {
        "node": ">=4"
      }
    },
    "node_modules/tr46": {
      "version": "0.0.3",
      "resolved": "https://registry.npmjs.org/tr46/-/tr46-0.0.3.tgz",
      "integrity": "sha512-N3WMsuqV66lT30CrXNbEjx4GEwlow3v6rr4mCcv6prnfwhS01rkgyFdjPNBYd9br7LpXV1+Emh01fHnq2Gdgrw==",
      "dev": true,
      "license": "MIT"
    },
    "node_modules/typescript": {
      "version": "7.0.2",
      "resolved": "https://registry.npmjs.org/typescript/-/typescript-7.0.2.tgz",
      "integrity": "sha512-8FYau96o3NKOhbjKi/qNvG/W5jhzxkbdm5sj9AbZ/5T5sWqn3hJgLfGx27sRKZWTvyzCP8dLRBTf5tBTSRVUNA==",
      "dev": true,
      "license": "Apache-2.0",
      "bin": {
        "tsc": "bin/tsc"
      },
      "engines": {
        "node": ">=16.20.0"
      },
      "optionalDependencies": {
        "@typescript/typescript-aix-ppc64": "7.0.2",
        "@typescript/typescript-darwin-arm64": "7.0.2",
        "@typescript/typescript-darwin-x64": "7.0.2",
        "@typescript/typescript-freebsd-arm64": "7.0.2",
        "@typescript/typescript-freebsd-x64": "7.0.2",
        "@typescript/typescript-linux-arm": "7.0.2",
        "@typescript/typescript-linux-arm64": "7.0.2",
        "@typescript/typescript-linux-loong64": "7.0.2",
        "@typescript/typescript-linux-mips64el": "7.0.2",
        "@typescript/typescript-linux-ppc64": "7.0.2",
        "@typescript/typescript-linux-riscv64": "7.0.2",
        "@typescript/typescript-linux-s390x": "7.0.2",
        "@typescript/typescript-linux-x64": "7.0.2",
        "@typescript/typescript-netbsd-arm64": "7.0.2",
        "@typescript/typescript-netbsd-x64": "7.0.2",
        "@typescript/typescript-openbsd-arm64": "7.0.2",
        "@typescript/typescript-openbsd-x64": "7.0.2",
        "@typescript/typescript-sunos-x64": "7.0.2",
        "@typescript/typescript-win32-arm64": "7.0.2",
        "@typescript/typescript-win32-x64": "7.0.2"
      }
    },
    "node_modules/undici-types": {
      "version": "8.9.0",
      "resolved": "https://registry.npmjs.org/undici-types/-/undici-types-8.9.0.tgz",
      "integrity": "sha512-KTDyRTYX8sWmKXAikPHHSyc63CRPETMctyjKFupcC6OBLXT3xsN0e9aF7m+mIXutFWpUXuedtowG7iLOzp0kQg==",
      "dev": true,
      "license": "MIT"
    },
    "node_modules/universalify": {
      "version": "2.0.1",
      "resolved": "https://registry.npmjs.org/universalify/-/universalify-2.0.1.tgz",
      "integrity": "sha512-gptHNQghINnc/vTGIk0SOFGFNXw7JVrlRUtConJRlvaw6DuX0wO5Jeko9sWrMBhh+PsYAZ7oXAiOnf/UKogyiw==",
      "dev": true,
      "license": "MIT",
      "engines": {
        "node": ">= 10.0.0"
      }
    },
    "node_modules/uuid": {
      "version": "14.0.3",
      "resolved": "https://registry.npmjs.org/uuid/-/uuid-14.0.3.tgz",
      "integrity": "sha512-Sc5u1VVln93y1Xnz4R+CunOZc7R9LUcTmhj7Rc7Eg3HVcf7rVn2FOoa812fe0W8p3jD7xeoGSYbfqCGGH2B6xw==",
      "dev": true,
      "funding": [
        "https://github.com/sponsors/broofa",
        "https://github.com/sponsors/ctavan"
      ],
      "license": "MIT",
      "bin": {
        "uuid": "dist-node/bin/uuid"
      }
    },
    "node_modules/web-streams-polyfill": {
      "version": "3.3.3",
      "resolved": "https://registry.npmjs.org/web-streams-polyfill/-/web-streams-polyfill-3.3.3.tgz",
      "integrity": "sha512-d2JWLCivmZYTSIoge9MsgFCZrt571BikcWGYkjC1khllbTeDlGqZ2D8vD8E/lJa8WGWbb7Plm8/XJYV7IJHZZw==",
      "dev": true,
      "license": "MIT",
      "engines": {
        "node": ">= 8"
      }
    },
    "node_modules/webidl-conversions": {
      "version": "3.0.1",
      "resolved": "https://registry.npmjs.org/webidl-conversions/-/webidl-conversions-3.0.1.tgz",
      "integrity": "sha512-2JAn3z8AR6rjK8Sm8orRC0h/bcl/DqL7tRPdGZ4I1CjdF+EaMLmYxBHyXuKL849eucPFhvBoxMsflfOb8kxaeQ==",
      "dev": true,
      "license": "BSD-2-Clause"
    },
    "node_modules/whatwg-url": {
      "version": "5.0.0",
      "resolved": "https://registry.npmjs.org/whatwg-url/-/whatwg-url-5.0.0.tgz",
      "integrity": "sha512-saE57nupxk6v3HY35+jzBwYa0rKSy0XR8JSxZPwgLr7ys0IBzhGviA1/TUGJLmSVqs8pb9AnvICXEuOHLprYTw==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "tr46": "~0.0.3",
        "webidl-conversions": "^3.0.0"
      }
    },
    "node_modules/ws": {
      "version": "8.22.0",
      "resolved": "https://registry.npmjs.org/ws/-/ws-8.22.0.tgz",
      "integrity": "sha512-Ydggc987+RO0AnWtZ/7Wq9FtNvcrL1b/RO0ud9mWjUPgDrsAAwQSF51sm2hm1XofbU/4jkpGEsLFsZZxU+1DOg==",
      "license": "MIT",
      "engines": {
        "node": ">=10.0.0"
      },
      "peerDependencies": {
        "bufferutil": "^4.0.1",
        "utf-8-validate": ">=5.0.2"
      },
      "peerDependenciesMeta": {
        "bufferutil": {
          "optional": true
        },
        "utf-8-validate": {
          "optional": true
        }
      }
    },
    "node_modules/xml2js": {
      "version": "0.5.0",
      "resolved": "https://registry.npmjs.org/xml2js/-/xml2js-0.5.0.tgz",
      "integrity": "sha512-drPFnkQJik/O+uPKpqSgr22mpuFHqKdbS835iAQrUC73L2F5WkboIRd63ai/2Yg6I1jzifPFKH2NTK+cfglkIA==",
      "dev": true,
      "license": "MIT",
      "dependencies": {
        "sax": ">=0.6.0",
        "xmlbuilder": "~11.0.0"
      },
      "engines": {
        "node": ">=4.0.0"
      }
    },
    "node_modules/xmlbuilder": {
      "version": "11.0.1",
      "resolved": "https://registry.npmjs.org/xmlbuilder/-/xmlbuilder-11.0.1.tgz",
      "integrity": "sha512-fDlsI/kFEx7gLvbecc0/ohLG50fugQp8ryHzMTuW9vSa1GJ0XYWKnhsUx7oie3G98+r56aTQIUB4kht42R3JvA==",
      "dev": true,
      "license": "MIT",
      "engines": {
        "node": ">=4.0"
      }
    }
  }
}
````
