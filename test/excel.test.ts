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

test("checkSheetName follows Excel's rules", async () => {
  const { checkSheetName } = await import("../src/excel.ts");
  assert.equal(checkSheetName("  TC 2026-11 ", ["Sheet1"]), "TC 2026-11");
  assert.throws(() => checkSheetName("", []), /empty/);
  assert.throws(() => checkSheetName("x".repeat(32), []), /31/);
  assert.throws(() => checkSheetName("a/b", []), /contains/);
  assert.throws(() => checkSheetName("'quoted'", []), /contains/);
  assert.throws(() => checkSheetName("History", []), /reserved/);
  assert.throws(() => checkSheetName("sheet1", ["Sheet1"]), /already exists/);
});

test("nextRow is the first row below the used range", async () => {
  const { nextRow } = await import("../src/excel.ts");
  assert.equal(nextRow(null), 1);
  assert.equal(nextRow(parseRange("A1:K120")), 121);
  assert.equal(nextRow(parseRange("C5:D9")), 10);
});

test("findInGrid reports addresses relative to the used range and stops at the limit", async () => {
  const { findInGrid } = await import("../src/excel.ts");
  const grid = [["TC-001", "login"], ["", "TC-002 logout"], [null, 42]];
  const has = (s: string) => s.toLowerCase().includes("tc-00");
  const hits = findInGrid("Tests", parseRange("B3:C5"), grid, has, 10);
  assert.deepEqual(hits.map((h) => h.address), ["B3", "C4"]);
  assert.equal(hits[0].hash, `sha256:${hash("TC-001")}`);
  assert.equal(findInGrid("Tests", parseRange("B3:C5"), grid, has, 1).length, 1);
  assert.deepEqual(findInGrid("Tests", parseRange("B3:C5"), grid, (s) => s === "42", 10).map((h) => h.address), ["C5"]);
});

test("formatPatch validates and splits font / format / numberFormat", async () => {
  const { formatPatch } = await import("../src/excel.ts");
  assert.deepEqual(formatPatch({ bold: true, fontColor: "#c00000", wrapText: true }), {
    font: { bold: true, color: "#C00000" },
    format: { wrapText: true },
    numberFormat: undefined,
  });
  assert.equal(formatPatch({ numberFormat: "@" }).numberFormat, "@");
  assert.throws(() => formatPatch({}), /Nothing to change/);
  assert.throws(() => formatPatch({ fontColor: "red" }), /fontColor/);
  assert.throws(() => formatPatch({ fontSize: 0 }), /fontSize/);
});
