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
