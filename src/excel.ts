// Excel-on-SharePoint operations with the guards shared test/spec workbooks need:
// workbook sessions (sessionless calls on big files time out), base checks before writing,
// read-back hash verification after writing, dry runs until send=true.
import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { HOME, graph, GraphError, resolveUrl, wb, type DriveItem, type GraphOptions } from "./graph.ts";

export const SEP = "-----update-----";
const BATCH_MAX = 20;
const SESSION_IDLE_MS = 4 * 60 * 1000; // Graph drops idle persistent sessions after ~5 min
// Workbook aliases: $GRAPH_EXCEL_HOME/workbooks.json (see workbooks.example.json).
const ALIAS_FILES = [path.join(HOME, "workbooks.json")];

/** A cell value as Graph returns it. */
export type Cell = string | number | boolean | null;
type Grid = Cell[][];

export interface Range {
  address: string;
  c1: number;
  r1: number;
  rows: number;
  cols: number;
}

// ---------- helpers ----------

export const hash = (v: unknown) => crypto.createHash("sha256").update(String(v ?? "")).digest("hex").slice(0, 12);
const normEol = (s: unknown) => String(s ?? "").replace(/\r\n/g, "\n");
const preview = (v: unknown, n = 80) => {
  const s = String(v ?? "").replace(/\r?\n/g, "⏎");
  return s.length > n ? `${s.slice(0, n)}… (${s.length} chars)` : s;
};

const colNum = (letters: string) => [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
const colName = (n: number) => {
  let s = "";
  for (; n > 0; n = Math.floor((n - 1) / 26)) s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
  return s;
};
const ADDR = /^([A-Z]{1,3})([1-9][0-9]{0,6})(?::([A-Z]{1,3})([1-9][0-9]{0,6}))?$/;

export function parseRange(address: unknown): Range {
  const a = String(address ?? "").trim().toUpperCase().replace(/\$/g, "");
  const m = a.match(ADDR);
  if (!m) throw new Error(`Invalid range "${address}" (use one contiguous range like H60 or H60:J60)`);
  const c1 = colNum(m[1]), r1 = Number(m[2]);
  const c2 = m[3] ? colNum(m[3]) : c1, r2 = m[4] ? Number(m[4]) : r1;
  if (c2 < c1 || r2 < r1) throw new Error(`Invalid range "${address}" (top-left first)`);
  return { address: a, c1, r1, rows: r2 - r1 + 1, cols: c2 - c1 + 1 };
}
const cellAddr = (rg: Range, i: number, j: number) => `${colName(rg.c1 + j)}${rg.r1 + i}`;

function to2d<T>(values: T | T[] | T[][], rg: Range, name: string): T[][] {
  let grid: T[][];
  if (!Array.isArray(values)) grid = [[values]];
  else if (!Array.isArray(values[0])) grid = rg.rows === 1 ? [values as T[]] : (values as T[]).map((v) => [v]);
  else grid = values as T[][];
  if (grid.length !== rg.rows || grid.some((r) => !Array.isArray(r) || r.length !== rg.cols))
    throw new Error(`${name} for ${rg.address} must be ${rg.rows}×${rg.cols}`);
  return grid;
}

/** Base check for one cell: exact text, "sha256:<prefix>" from read_range, or null/undefined = don't check. */
export type Expected = string | null | undefined;

function matches(current: Cell, expected: Expected): boolean {
  if (expected === null || expected === undefined) return true;
  const e = String(expected);
  if (e.startsWith("sha256:")) {
    const want = e.slice(7);
    return want.length >= 8 && (hash(current).startsWith(want) || hash(normEol(current)).startsWith(want));
  }
  return String(current ?? "") === e || normEol(current) === normEol(e);
}

// ---------- workbook + session ----------

interface Alias {
  name?: string;
  driveId: string;
  itemId: string;
}

interface BatchRequest {
  method: string;
  url: string;
  body?: unknown;
}

interface BatchResponse {
  id?: string;
  status: number;
  body: any;
}

interface Workbook {
  it: DriveItem;
  base: string;
  sheetUrl(sheet: string): string;
  rangeUrl(sheet: string, address: string): string;
  call(method: string, url: string, body?: unknown, opts?: GraphOptions): Promise<any>;
  batch(requests: BatchRequest[]): Promise<BatchResponse[]>;
  ok(r: BatchResponse, what: string): any;
}

const items = new Map<string, DriveItem>(); // file ref -> driveItem
const sessions = new Map<string, { id: string; last: number }>(); // workbook path -> session

function aliases(): Record<string, Alias> {
  const all: Record<string, Alias> = {};
  for (const f of ALIAS_FILES) {
    try {
      Object.assign(all, JSON.parse(fs.readFileSync(f, "utf8")));
    } catch {
      // no alias file
    }
  }
  return all;
}

async function item(file: unknown): Promise<DriveItem> {
  const ref = String(file ?? "").trim();
  if (!ref) throw new Error("Missing file (SharePoint URL, alias from ~/.graph-excel/workbooks.json, or driveId:itemId)");
  const cached = items.get(ref);
  if (cached) return cached;
  let it: DriveItem;
  const alias = aliases()[ref];
  if (alias) it = { id: alias.itemId, name: alias.name, parentReference: { driveId: alias.driveId } };
  else if (/^https:\/\//.test(ref)) it = await resolveUrl(ref);
  else if (/^b![^:]+:[A-Z0-9]+$/.test(ref)) {
    const [driveId, id] = ref.split(":");
    it = { id, parentReference: { driveId } };
  } else throw new Error(`Unknown workbook "${ref}" — pass the SharePoint URL or an alias: ${Object.keys(aliases()).join(", ") || "(none)"}`);
  items.set(ref, it);
  return it;
}

async function session(base: string, { fresh = false } = {}): Promise<string> {
  const s = sessions.get(base);
  if (!fresh && s && Date.now() - s.last < SESSION_IDLE_MS) {
    s.last = Date.now();
    return s.id;
  }
  const { id } = await graph("POST", `${base}/createSession`, { persistChanges: true }, { retries: 2 });
  sessions.set(base, { id, last: Date.now() });
  return id;
}

const sessionLost = (e: unknown) => e instanceof GraphError && /session/i.test(e.message) && [400, 404, 410].includes(e.status);

// Workbook context: call(method, relUrl, body) and batch(requests) inside a session, re-opened once if it expired.
async function open(file: unknown): Promise<Workbook> {
  const it = await item(file);
  const base = wb(it);
  const withSession = async <T>(fn: (sid: string) => Promise<T>): Promise<T> => {
    try {
      return await fn(await session(base));
    } catch (e) {
      if (!sessionLost(e)) throw e;
      return fn(await session(base, { fresh: true }));
    }
  };
  const sheetUrl = (sheet: string) => {
    if (!sheet) throw new Error("Missing sheet name");
    return `${base}/worksheets('${encodeURIComponent(String(sheet).replace(/'/g, "''"))}')`;
  };
  const rangeUrl = (sheet: string, address: string) => `${sheetUrl(sheet)}/range(address='${address}')`;
  const call = (method: string, url: string, body?: unknown, opts: GraphOptions = {}) =>
    withSession((sid) => graph(method, url, body, { ...opts, session: sid }));

  // requests -> responses in the same order. Retryable read failures are retried one by one; writes are never
  // blindly retried here (callers verify by reading back).
  async function batch(requests: BatchRequest[]): Promise<BatchResponse[]> {
    const out: BatchResponse[] = [];
    for (let i = 0; i < requests.length; i += BATCH_MAX) {
      const chunk = requests.slice(i, i + BATCH_MAX);
      const res = await withSession((sid) =>
        graph(
          "POST",
          "/$batch",
          {
            requests: chunk.map((r, k) => ({
              id: String(k + 1),
              method: r.method,
              url: r.url,
              headers: { "Content-Type": "application/json", "workbook-session-id": sid },
              ...(r.body === undefined ? {} : { body: r.body }),
            })),
          },
          { retries: 1 },
        ),
      );
      const byId = new Map<string, BatchResponse>(res.responses.map((r: BatchResponse) => [r.id, r]));
      for (let k = 0; k < chunk.length; k++) {
        let r = byId.get(String(k + 1)) ?? { status: 500, body: { error: { message: "missing batch response" } } };
        if (r.status >= 400 && chunk[k].method === "GET" && [429, 502, 503, 504].includes(r.status)) {
          try {
            r = { status: 200, body: await call("GET", chunk[k].url) };
          } catch (e) {
            r = { status: e instanceof GraphError ? e.status : 500, body: { error: { message: (e as Error).message } } };
          }
        }
        out.push(r);
      }
    }
    return out;
  }
  const ok = (r: BatchResponse, what: string) => {
    if (r.status >= 400) throw new Error(`${what}: ${r.status} ${JSON.stringify(r.body?.error ?? r.body).slice(0, 300)}`);
    return r.body;
  };
  return { it, base, sheetUrl, rangeUrl, call, batch, ok };
}

async function readValues(wbx: Workbook, sheet: string, ranges: Range[]): Promise<Grid[]> {
  const res = await wbx.batch(ranges.map((rg) => ({ method: "GET", url: `${wbx.rangeUrl(sheet, rg.address)}?$select=address,values` })));
  return res.map((r, i) => wbx.ok(r, `read ${ranges[i].address}`).values);
}

async function readFills(wbx: Workbook, sheet: string, addresses: string[]): Promise<(string | null)[]> {
  const res = await wbx.batch(addresses.map((a) => ({ method: "GET", url: `${wbx.rangeUrl(sheet, a)}/format/fill` })));
  return res.map((r, i) => wbx.ok(r, `read fill ${addresses[i]}`).color || null);
}

// ---------- operations ----------

export async function resolve({ file }: { file: string }) {
  const wbx = await open(file);
  const full = await graph("GET", `/drives/${wbx.it.parentReference.driveId}/items/${wbx.it.id}?$select=name,webUrl,lastModifiedDateTime,lastModifiedBy`);
  const ws: { value: Array<{ name: string; position: number; visibility: string }> } = await wbx.call(
    "GET",
    `${wbx.base}/worksheets?$select=name,position,visibility`,
  );
  const used = await wbx.batch(ws.value.map((w) => ({ method: "GET", url: `${wbx.sheetUrl(w.name)}/usedRange(valuesOnly=true)?$select=address` })));
  return {
    name: full.name,
    driveId: wbx.it.parentReference.driveId,
    itemId: wbx.it.id,
    webUrl: full.webUrl,
    lastModified: `${full.lastModifiedDateTime} by ${full.lastModifiedBy?.user?.displayName ?? "?"}`,
    sheets: ws.value.map((w, i) => ({
      position: w.position,
      name: w.name,
      ...(w.visibility !== "Visible" ? { visibility: w.visibility } : {}),
      usedRange: used[i].status < 400 ? String(used[i].body.address).replace(/^.*!/, "") : null,
    })),
  };
}

interface ReadCell {
  address: string;
  hash: string;
  value?: Cell;
  length?: number;
  fill?: string | null;
}

export async function readRange({ file, sheet, range, fills = false, hashOnly = false }: { file: string; sheet: string; range: string; fills?: boolean; hashOnly?: boolean }) {
  const rg = parseRange(range);
  const wbx = await open(file);
  const [values] = await readValues(wbx, sheet, [rg]);
  const cells: ReadCell[] = [];
  values.forEach((row, i) =>
    row.forEach((v, j) =>
      cells.push({ address: cellAddr(rg, i, j), hash: `sha256:${hash(v)}`, ...(hashOnly ? { length: String(v).length } : { value: v }) }),
    ),
  );
  if (fills) {
    if (cells.length > 400) throw new Error("fills=true is limited to 400 cells per call");
    const colors = await readFills(wbx, sheet, cells.map((c) => c.address));
    cells.forEach((c, k) => (c.fill = colors[k]));
  }
  return { sheet, range: rg.address, cells };
}

export interface WriteBlock {
  range: string;
  values: Cell | Cell[] | Cell[][];
  expected?: Expected | Expected[] | Expected[][];
}

interface BlockReport {
  range: string;
  status: "conflict" | "will-write" | "unchanged" | "written-verified" | "written-mismatch";
  conflicts: Array<{ address: string; current: string; currentHash: string }>;
  changes: Array<{ address: string; from: string; to: string }>;
  mismatches?: Array<{ address: string; wanted: string; got: string }>;
  note?: string;
  error?: string;
  hashes?: string[];
}

// Each block is one contiguous range. A block whose current content does not match `expected` is skipped as a
// conflict (never partially written).
export async function writeValues({ file, sheet, writes, send = false }: { file: string; sheet: string; writes: WriteBlock[]; send?: boolean }) {
  if (!Array.isArray(writes) || !writes.length) throw new Error("writes must be a non-empty array");
  const plan = writes.map((w) => {
    const rg = parseRange(w.range);
    const values = to2d<Cell>(w.values, rg, "values").map((row) => row.map((v) => (typeof v === "string" ? normEol(v) : v)));
    return { rg, values, expected: w.expected === undefined ? null : to2d<Expected>(w.expected, rg, "expected") };
  });
  const wbx = await open(file);
  const current = await readValues(wbx, sheet, plan.map((p) => p.rg));

  const report: BlockReport[] = plan.map((p, k) => {
    const conflicts: BlockReport["conflicts"] = [];
    const changes: BlockReport["changes"] = [];
    p.values.forEach((row, i) =>
      row.forEach((v, j) => {
        const cur = current[k][i][j];
        const addr = cellAddr(p.rg, i, j);
        if (p.expected && !matches(cur, p.expected[i][j])) conflicts.push({ address: addr, current: preview(cur), currentHash: `sha256:${hash(cur)}` });
        if (String(cur ?? "") !== String(v ?? "")) changes.push({ address: addr, from: preview(cur), to: preview(v) });
      }),
    );
    return { range: p.rg.address, status: conflicts.length ? "conflict" : changes.length ? "will-write" : "unchanged", conflicts, changes };
  });
  if (!send) return { dryRun: true, sheet, blocks: report, note: "Nothing written. Call again with send=true after the user approved." };

  const todo = plan.map((p, k) => ({ p, r: report[k] })).filter(({ r }) => r.status === "will-write");
  const res = await wbx.batch(todo.map(({ p }) => ({ method: "PATCH", url: wbx.rangeUrl(sheet, p.rg.address), body: { values: p.values } })));
  // Read back everything (a 504 may still have applied the write), retry once what didn't land.
  let after = await readValues(wbx, sheet, todo.map(({ p }) => p.rg));
  const landed = (k: number) => todo[k].p.values.every((row, i) => row.every((v, j) => normEol(after[k][i][j]) === normEol(v)));
  const retry = todo.map((_, k) => k).filter((k) => !landed(k) && res[k].status >= 500);
  if (retry.length) {
    await wbx.batch(retry.map((k) => ({ method: "PATCH", url: wbx.rangeUrl(sheet, todo[k].p.rg.address), body: { values: todo[k].p.values } })));
    after = await readValues(wbx, sheet, todo.map(({ p }) => p.rg));
  }
  todo.forEach(({ p, r }, k) => {
    const bad: NonNullable<BlockReport["mismatches"]> = [];
    p.values.forEach((row, i) =>
      row.forEach((v, j) => {
        const got = after[k][i][j];
        if (String(got) === String(v) || normEol(got) === normEol(v)) return;
        bad.push({ address: cellAddr(p.rg, i, j), wanted: preview(v), got: preview(got) });
      }),
    );
    r.status = bad.length ? "written-mismatch" : "written-verified";
    if (bad.length) r.mismatches = bad;
    if (res[k].status >= 400 && !bad.length) r.note = `request returned ${res[k].status} but the read-back matches`;
    if (res[k].status >= 400 && bad.length) r.error = JSON.stringify(res[k].body?.error ?? res[k].body).slice(0, 300);
    r.hashes = after[k].flat().map((v) => `sha256:${hash(v)}`);
  });
  return { dryRun: false, sheet, blocks: report };
}

export function composeUpdate(cur: Cell, text: string, mode: "replace" | "append" = "replace"): { head: string; next: string } {
  const s = String(cur ?? "");
  if (!s.trim()) throw new Error("cell is empty — there is no original part to keep; use write_values");
  if (s.includes("\r\n")) return composeUpdate(normEol(s), text, mode);
  // Graph returns line breaks as \n and Excel stores a written \n as \r\n; a written \r\n would be stored
  // with a stray _x000D_, so always compose with \n.
  const eol = "\n";
  const body = normEol(text).replace(/\s+$/, "");
  const idx = s.indexOf(SEP);
  if (idx < 0) {
    const head = s.replace(/\s+$/, "") + eol + eol + SEP;
    return { head, next: head + eol + body };
  }
  const head = s.slice(0, idx + SEP.length);
  if (mode === "append") {
    const old = s.slice(head.length).replace(/^\s+/, "").replace(/\s+$/, "");
    return { head, next: head + eol + (old ? old + eol : "") + body };
  }
  return { head, next: head + eol + body };
}

export interface AppendUpdateArgs {
  file: string;
  sheet: string;
  row: number;
  /** Column letter -> text of the update part, e.g. { H: "...", J: "..." }. */
  update: Record<string, string>;
  expected?: Record<string, string>;
  mode?: "replace" | "append";
  fill?: string | null;
  fillRange?: string;
  send?: boolean;
}

type PlannedCell =
  | { address: string; status: "conflict"; current: string; currentHash: string }
  | { address: string; status: "will-write" | "unchanged"; head: string; next: string; cur: Cell };

export async function appendUpdate({ file, sheet, row, update, expected = {}, mode = "replace", fill = "#FFFF00", fillRange, send = false }: AppendUpdateArgs) {
  const r = Number(row);
  if (!Number.isInteger(r) || r < 1) throw new Error("row must be a positive integer");
  const keys = Object.keys(update ?? {});
  const cols = keys.map((c) => c.toUpperCase());
  if (!cols.length || cols.some((c) => !/^[A-Z]{1,3}$/.test(c))) throw new Error('update must map column letters to text, e.g. {"H": "..."}');
  if (!["replace", "append"].includes(mode)) throw new Error('mode must be "replace" or "append"');
  const wbx = await open(file);
  const current = (await readValues(wbx, sheet, cols.map((c) => parseRange(`${c}${r}`)))).map((v) => v[0][0]);
  const exp = Object.fromEntries(Object.entries(expected ?? {}).map(([k, v]) => [k.toUpperCase(), v]));

  const cells: PlannedCell[] = cols.map((c, k) => {
    const cur = current[k];
    const address = `${c}${r}`;
    if (exp[c] !== undefined && !matches(cur, exp[c])) return { address, status: "conflict", current: preview(cur), currentHash: `sha256:${hash(cur)}` };
    const { head, next } = composeUpdate(cur, update[keys[k]], mode);
    return { address, status: next === String(cur) ? "unchanged" : "will-write", head, next, cur };
  });
  const fr = parseRange(fillRange || `${colName(Math.min(...cols.map(colNum)))}${r}:${colName(Math.max(...cols.map(colNum)))}${r}`);
  const shown = () =>
    cells.map((c) =>
      c.status === "conflict"
        ? c
        : { address: c.address, status: c.status, from: preview(c.cur, 60), updatePart: preview(c.next.slice(c.next.indexOf(SEP)), 300) },
    );
  if (!send) {
    const [fillNow] = fill ? await readFills(wbx, sheet, [fr.address]) : [null];
    return {
      dryRun: true,
      sheet,
      row: r,
      cells: shown(),
      fill: fill ? { range: fr.address, from: fillNow, to: fill } : null,
      note: "Nothing written. Call again with send=true after the user approved.",
    };
  }
  if (cells.some((c) => c.status === "conflict"))
    return { dryRun: false, written: false, sheet, row: r, cells: shown(), note: "Conflict: someone changed a cell since the base was read — nothing written." };

  const planned = cells as Exclude<PlannedCell, { status: "conflict" }>[];
  const todo = planned.filter((c) => c.status === "will-write");
  const result = todo.length
    ? await writeValues({ file, sheet, send: true, writes: todo.map((c) => ({ range: c.address, values: c.next, expected: c.cur === null ? null : String(c.cur) })) })
    : { blocks: [] as BlockReport[] };
  const after = (await readValues(wbx, sheet, planned.map((c) => parseRange(c.address)))).map((v) => v[0][0]);
  const out = planned.map((c, k) => {
    const block = result.blocks.find((b) => b.range === c.address);
    return {
      address: c.address,
      status: block ? block.status : c.status,
      originalKept: normEol(after[k]).startsWith(normEol(c.head)),
      hash: `sha256:${hash(after[k])}`,
      ...(block?.conflicts.length ? { conflicts: block.conflicts } : {}),
      ...(block?.mismatches ? { mismatches: block.mismatches } : {}),
    };
  });
  let fillResult = null;
  if (fill && out.every((c) => c.originalKept && c.status !== "conflict")) {
    fillResult = (await setFill({ file, sheet, ranges: [fr.address], color: fill, send: true })).ranges[0];
  }
  return { dryRun: false, sheet, row: r, cells: out, fill: fillResult };
}

interface FillRow {
  range: string;
  from: string;
  to: string;
  now?: string;
  status?: "verified" | "mismatch";
}

/** color: "#FFFF00" etc., or null to clear the fill. */
export async function setFill({ file, sheet, ranges, color = null, send = false }: { file: string; sheet: string; ranges: string[]; color?: string | null; send?: boolean }) {
  if (!Array.isArray(ranges) || !ranges.length) throw new Error("ranges must be a non-empty array of contiguous ranges");
  const rgs = ranges.map(parseRange);
  if (color !== null && !/^#[0-9A-Fa-f]{6}$/.test(String(color))) throw new Error('color must be like "#FFFF00", or null to clear');
  const want = color === null ? null : String(color).toUpperCase();
  const wbx = await open(file);
  const before = await readFills(wbx, sheet, rgs.map((r) => r.address));
  const rows: FillRow[] = rgs.map((r, k) => ({ range: r.address, from: before[k] ?? "(none or mixed)", to: want ?? "(cleared)" }));
  if (!send) return { dryRun: true, sheet, ranges: rows, note: "Nothing written. Call again with send=true after the user approved." };
  await wbx.batch(
    rgs.map((r) =>
      want === null
        ? { method: "POST", url: `${wbx.rangeUrl(sheet, r.address)}/format/fill/clear`, body: {} }
        : { method: "PATCH", url: `${wbx.rangeUrl(sheet, r.address)}/format/fill`, body: { color: want } },
    ),
  );
  const after = await readFills(wbx, sheet, rgs.map((r) => r.address));
  rows.forEach((row, k) => {
    const got = after[k] ? after[k].toUpperCase() : null;
    row.now = got ?? "(none)";
    row.status = got === want ? "verified" : "mismatch";
  });
  return { dryRun: false, sheet, ranges: rows };
}
