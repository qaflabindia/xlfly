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
