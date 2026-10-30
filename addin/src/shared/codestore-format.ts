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
