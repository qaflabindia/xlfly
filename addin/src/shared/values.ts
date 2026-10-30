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
