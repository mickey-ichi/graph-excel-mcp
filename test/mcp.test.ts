// MCP handshake over stdio, against both the .ts source and the built dist/ (no Graph calls).
import { test } from "node:test";
import assert from "node:assert/strict";
import { spawn } from "node:child_process";
import fs from "node:fs";
import path from "node:path";

const ROOT = path.join(import.meta.dirname, "..");
const TOOL_NAMES = [
  "login_status", "resolve", "read_range", "write_values", "append_update", "set_fill",
  "find", "add_sheet", "rename_sheet", "insert_rows", "append_rows", "set_format",
];

async function handshake(entry: string) {
  const p = spawn(process.execPath, [entry], { stdio: ["pipe", "pipe", "inherit"] });
  const replies: any[] = [];
  let buf = "";
  const done = new Promise<void>((resolve) =>
    p.stdout.on("data", (d) => {
      buf += d;
      let nl: number;
      while ((nl = buf.indexOf("\n")) >= 0) {
        replies.push(JSON.parse(buf.slice(0, nl)));
        buf = buf.slice(nl + 1);
      }
      if (replies.length === 3) resolve();
    }),
  );
  const rpc = (id: number, method: string, params?: unknown) => p.stdin.write(JSON.stringify({ jsonrpc: "2.0", id, method, params }) + "\n");
  rpc(1, "initialize", { protocolVersion: "2025-06-18" });
  rpc(2, "tools/list");
  rpc(3, "nope");
  await done;
  p.kill();
  return replies;
}

for (const entry of ["src/cli.ts", "dist/cli.js"]) {
  test(`${entry}: initialize, tools/list, unknown method`, { skip: !fs.existsSync(path.join(ROOT, entry)) && "not built" }, async () => {
    const [init, list, unknown] = await handshake(path.join(ROOT, entry));
    const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, "package.json"), "utf8"));
    assert.equal(init.result.protocolVersion, "2025-06-18");
    assert.deepEqual(init.result.serverInfo, { name: "graph-excel", version: pkg.version });
    assert.deepEqual(list.result.tools.map((t: { name: string }) => t.name), TOOL_NAMES);
    assert.equal(unknown.error.code, -32601);
  });
}
