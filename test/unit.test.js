// Offline tests: range parsing, update composition, and the MCP handshake (no Graph calls).
const test = require("node:test");
const assert = require("node:assert");
const path = require("node:path");
const { spawn } = require("node:child_process");
const excel = require("../excel");

test("parseRange accepts one contiguous range only", () => {
  assert.deepStrictEqual(excel.parseRange("h60:j61"), { address: "H60:J61", c1: 8, r1: 60, rows: 2, cols: 3 });
  assert.deepStrictEqual(excel.parseRange("$AF$1"), { address: "AF1", c1: 32, r1: 1, rows: 1, cols: 1 });
  assert.throws(() => excel.parseRange("H60:J60,H63:J63"));
  assert.throws(() => excel.parseRange("J60:H60"));
});

test("composeUpdate keeps the original part and always uses \\n", () => {
  let r = excel.composeUpdate("1. a\r\n2. b\r\n", "x\r\ny", "replace");
  assert.strictEqual(r.next, "1. a\n2. b\n\n-----update-----\nx\ny");
  r = excel.composeUpdate(r.next, "z", "replace");
  assert.strictEqual(r.next, "1. a\n2. b\n\n-----update-----\nz");
  r = excel.composeUpdate(r.next, "w", "append");
  assert.strictEqual(r.next, "1. a\n2. b\n\n-----update-----\nz\nw");
  assert.ok(r.next.startsWith(r.head));
  assert.throws(() => excel.composeUpdate("", "x"));
});

test("MCP server answers initialize and lists the six tools", async () => {
  const p = spawn(process.execPath, [path.join(__dirname, "..", "mcp.js")], { stdio: ["pipe", "pipe", "inherit"] });
  const replies = [];
  let buf = "";
  const done = new Promise((resolve) =>
    p.stdout.on("data", (d) => {
      buf += d;
      let nl;
      while ((nl = buf.indexOf("\n")) >= 0) {
        replies.push(JSON.parse(buf.slice(0, nl)));
        buf = buf.slice(nl + 1);
      }
      if (replies.length === 2) resolve();
    }),
  );
  p.stdin.write(JSON.stringify({ jsonrpc: "2.0", id: 1, method: "initialize", params: { protocolVersion: "2025-06-18" } }) + "\n");
  p.stdin.write(JSON.stringify({ jsonrpc: "2.0", id: 2, method: "tools/list" }) + "\n");
  await done;
  p.kill();
  assert.strictEqual(replies[0].result.serverInfo.name, "graph-excel");
  assert.deepStrictEqual(
    replies[1].result.tools.map((t) => t.name),
    ["login_status", "resolve", "read_range", "write_values", "append_update", "set_fill"],
  );
});
