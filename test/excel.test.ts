// Offline tests: range parsing and update composition (no Graph calls).
import { test } from "node:test";
import assert from "node:assert/strict";
import { composeUpdate, hash, parseRange } from "../src/excel.ts";

test("parseRange accepts one contiguous range only", () => {
  assert.deepEqual(parseRange("h60:j61"), { address: "H60:J61", c1: 8, r1: 60, rows: 2, cols: 3 });
  assert.deepEqual(parseRange("$AF$1"), { address: "AF1", c1: 32, r1: 1, rows: 1, cols: 1 });
  assert.throws(() => parseRange("H60:J60,H63:J63"));
  assert.throws(() => parseRange("J60:H60"));
  assert.throws(() => parseRange(""));
});

test("composeUpdate adds the separator once and keeps the original part", () => {
  const r = composeUpdate("1. a\n2. b\n", "x\ny");
  assert.equal(r.next, "1. a\n2. b\n\n-----update-----\nx\ny");
  assert.ok(r.next.startsWith(r.head));
});

test("composeUpdate replaces or appends after an existing separator", () => {
  const base = "1. a\n\n-----update-----\nold";
  assert.equal(composeUpdate(base, "new").next, "1. a\n\n-----update-----\nnew");
  assert.equal(composeUpdate(base, "more", "append").next, "1. a\n\n-----update-----\nold\nmore");
});

test("composeUpdate always writes \\n (a written \\r\\n would store _x000D_)", () => {
  const r = composeUpdate("1. a\r\n2. b\r\n", "x\r\ny");
  assert.equal(r.next, "1. a\n2. b\n\n-----update-----\nx\ny");
  assert.ok(!r.next.includes("\r"));
});

test("composeUpdate refuses an empty cell", () => {
  assert.throws(() => composeUpdate("", "x"));
  assert.throws(() => composeUpdate(null, "x"));
});

test("hash is a 12-char sha256 prefix of the text", () => {
  assert.equal(hash(""), "e3b0c44298fc");
  assert.equal(hash(null), hash(""));
  assert.equal(hash("a"), "ca978112ca1b");
});
