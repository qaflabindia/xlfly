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
