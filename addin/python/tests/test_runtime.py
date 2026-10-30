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
