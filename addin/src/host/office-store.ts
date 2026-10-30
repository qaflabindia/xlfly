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
