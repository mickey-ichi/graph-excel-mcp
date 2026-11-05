// Excel-on-SharePoint operations with the guards shared test/spec workbooks need:
// workbook sessions (sessionless calls on big files time out), base checks before writing,
// read-back hash verification after writing, dry runs until send=true.
import crypto from "node:crypto";
import fs from "node:fs";
import path from "node:path";
import { HOME, graph, GraphError, resolveUrl, wb } from "./graph.js";
export const SEP = "-----update-----";
const BATCH_MAX = 20;
const SESSION_IDLE_MS = 4 * 60 * 1000; // Graph drops idle persistent sessions after ~5 min
// Workbook aliases: $GRAPH_EXCEL_HOME/workbooks.json (see workbooks.example.json).
const ALIAS_FILES = [path.join(HOME, "workbooks.json")];
// ---------- helpers ----------
export const hash = (v) => crypto.createHash("sha256").update(String(v ?? "")).digest("hex").slice(0, 12);
const normEol = (s) => String(s ?? "").replace(/\r\n/g, "\n");
const preview = (v, n = 80) => {
    const s = String(v ?? "").replace(/\r?\n/g, "⏎");
    return s.length > n ? `${s.slice(0, n)}… (${s.length} chars)` : s;
};
const colNum = (letters) => [...letters].reduce((n, c) => n * 26 + c.charCodeAt(0) - 64, 0);
const colName = (n) => {
    let s = "";
    for (; n > 0; n = Math.floor((n - 1) / 26))
        s = String.fromCharCode(65 + ((n - 1) % 26)) + s;
    return s;
};
const ADDR = /^([A-Z]{1,3})([1-9][0-9]{0,6})(?::([A-Z]{1,3})([1-9][0-9]{0,6}))?$/;
export function parseRange(address) {
    const a = String(address ?? "").trim().toUpperCase().replace(/\$/g, "");
    const m = a.match(ADDR);
    if (!m)
        throw new Error(`Invalid range "${address}" (use one contiguous range like H60 or H60:J60)`);
    const c1 = colNum(m[1]), r1 = Number(m[2]);
    const c2 = m[3] ? colNum(m[3]) : c1, r2 = m[4] ? Number(m[4]) : r1;
    if (c2 < c1 || r2 < r1)
        throw new Error(`Invalid range "${address}" (top-left first)`);
    return { address: a, c1, r1, rows: r2 - r1 + 1, cols: c2 - c1 + 1 };
}
const cellAddr = (rg, i, j) => `${colName(rg.c1 + j)}${rg.r1 + i}`;
function to2d(values, rg, name) {
    let grid;
    if (!Array.isArray(values))
        grid = [[values]];
    else if (!Array.isArray(values[0]))
        grid = rg.rows === 1 ? [values] : values.map((v) => [v]);
    else
        grid = values;
    if (grid.length !== rg.rows || grid.some((r) => !Array.isArray(r) || r.length !== rg.cols))
        throw new Error(`${name} for ${rg.address} must be ${rg.rows}×${rg.cols}`);
    return grid;
}
function matches(current, expected) {
    if (expected === null || expected === undefined)
        return true;
    const e = String(expected);
    if (e.startsWith("sha256:")) {
        const want = e.slice(7);
        return want.length >= 8 && (hash(current).startsWith(want) || hash(normEol(current)).startsWith(want));
    }
    return String(current ?? "") === e || normEol(current) === normEol(e);
}
const items = new Map(); // file ref -> driveItem
const sessions = new Map(); // workbook path -> session
function aliases() {
    const all = {};
    for (const f of ALIAS_FILES) {
        try {
            Object.assign(all, JSON.parse(fs.readFileSync(f, "utf8")));
        }
        catch {
            // no alias file
        }
    }
    return all;
}
async function item(file) {
    const ref = String(file ?? "").trim();
    if (!ref)
        throw new Error("Missing file (SharePoint URL, alias from ~/.graph-excel/workbooks.json, or driveId:itemId)");
    const cached = items.get(ref);
    if (cached)
        return cached;
    let it;
    const alias = aliases()[ref];
    if (alias)
        it = { id: alias.itemId, name: alias.name, parentReference: { driveId: alias.driveId } };
    else if (/^https:\/\//.test(ref))
        it = await resolveUrl(ref);
    else if (/^b![^:]+:[A-Z0-9]+$/.test(ref)) {
        const [driveId, id] = ref.split(":");
        it = { id, parentReference: { driveId } };
    }
    else
        throw new Error(`Unknown workbook "${ref}" — pass the SharePoint URL or an alias: ${Object.keys(aliases()).join(", ") || "(none)"}`);
    items.set(ref, it);
    return it;
}
async function session(base, { fresh = false } = {}) {
    const s = sessions.get(base);
    if (!fresh && s && Date.now() - s.last < SESSION_IDLE_MS) {
        s.last = Date.now();
        return s.id;
    }
    const { id } = await graph("POST", `${base}/createSession`, { persistChanges: true }, { retries: 2 });
    sessions.set(base, { id, last: Date.now() });
    return id;
}
const sessionLost = (e) => e instanceof GraphError && /session/i.test(e.message) && [400, 404, 410].includes(e.status);
// Workbook context: call(method, relUrl, body) and batch(requests) inside a session, re-opened once if it expired.
async function open(file) {
    const it = await item(file);
    const base = wb(it);
    const withSession = async (fn) => {
        try {
            return await fn(await session(base));
        }
        catch (e) {
            if (!sessionLost(e))
                throw e;
            return fn(await session(base, { fresh: true }));
        }
    };
    const sheetUrl = (sheet) => {
        if (!sheet)
            throw new Error("Missing sheet name");
        return `${base}/worksheets('${encodeURIComponent(String(sheet).replace(/'/g, "''"))}')`;
    };
    const rangeUrl = (sheet, address) => `${sheetUrl(sheet)}/range(address='${address}')`;
    const call = (method, url, body, opts = {}) => withSession((sid) => graph(method, url, body, { ...opts, session: sid }));
    // requests -> responses in the same order. Retryable read failures are retried one by one; writes are never
    // blindly retried here (callers verify by reading back).
    async function batch(requests) {
        const out = [];
        for (let i = 0; i < requests.length; i += BATCH_MAX) {
            const chunk = requests.slice(i, i + BATCH_MAX);
            const res = await withSession((sid) => graph("POST", "/$batch", {
                requests: chunk.map((r, k) => ({
                    id: String(k + 1),
                    method: r.method,
                    url: r.url,
                    headers: { "Content-Type": "application/json", "workbook-session-id": sid },
                    ...(r.body === undefined ? {} : { body: r.body }),
                })),
            }, { retries: 1 }));
            const byId = new Map(res.responses.map((r) => [r.id, r]));
            for (let k = 0; k < chunk.length; k++) {
                let r = byId.get(String(k + 1)) ?? { status: 500, body: { error: { message: "missing batch response" } } };
                if (r.status >= 400 && chunk[k].method === "GET" && [429, 502, 503, 504].includes(r.status)) {
                    try {
                        r = { status: 200, body: await call("GET", chunk[k].url) };
                    }
                    catch (e) {
                        r = { status: e instanceof GraphError ? e.status : 500, body: { error: { message: e.message } } };
                    }
                }
                out.push(r);
            }
        }
        return out;
    }
    const ok = (r, what) => {
        if (r.status >= 400)
            throw new Error(`${what}: ${r.status} ${JSON.stringify(r.body?.error ?? r.body).slice(0, 300)}`);
        return r.body;
    };
    return { it, base, sheetUrl, rangeUrl, call, batch, ok };
}
async function readValues(wbx, sheet, ranges) {
    const res = await wbx.batch(ranges.map((rg) => ({ method: "GET", url: `${wbx.rangeUrl(sheet, rg.address)}?$select=address,values` })));
    return res.map((r, i) => wbx.ok(r, `read ${ranges[i].address}`).values);
}
async function readFills(wbx, sheet, addresses) {
    const res = await wbx.batch(addresses.map((a) => ({ method: "GET", url: `${wbx.rangeUrl(sheet, a)}/format/fill` })));
    return res.map((r, i) => wbx.ok(r, `read fill ${addresses[i]}`).color || null);
}
// ---------- operations ----------
export async function resolve({ file }) {
    const wbx = await open(file);
    const full = await graph("GET", `/drives/${wbx.it.parentReference.driveId}/items/${wbx.it.id}?$select=name,webUrl,lastModifiedDateTime,lastModifiedBy`);
    const ws = await wbx.call("GET", `${wbx.base}/worksheets?$select=name,position,visibility`);
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
export async function readRange({ file, sheet, range, fills = false, hashOnly = false }) {
    const rg = parseRange(range);
    const wbx = await open(file);
    const [values] = await readValues(wbx, sheet, [rg]);
    const cells = [];
    values.forEach((row, i) => row.forEach((v, j) => cells.push({ address: cellAddr(rg, i, j), hash: `sha256:${hash(v)}`, ...(hashOnly ? { length: String(v).length } : { value: v }) })));
    if (fills) {
        if (cells.length > 400)
            throw new Error("fills=true is limited to 400 cells per call");
        const colors = await readFills(wbx, sheet, cells.map((c) => c.address));
        cells.forEach((c, k) => (c.fill = colors[k]));
    }
    return { sheet, range: rg.address, cells };
}
// Each block is one contiguous range. A block whose current content does not match `expected` is skipped as a
// conflict (never partially written).
export async function writeValues({ file, sheet, writes, send = false }) {
    if (!Array.isArray(writes) || !writes.length)
        throw new Error("writes must be a non-empty array");
    const plan = writes.map((w) => {
        const rg = parseRange(w.range);
        const values = to2d(w.values, rg, "values").map((row) => row.map((v) => (typeof v === "string" ? normEol(v) : v)));
        return { rg, values, expected: w.expected === undefined ? null : to2d(w.expected, rg, "expected") };
    });
    const wbx = await open(file);
    const current = await readValues(wbx, sheet, plan.map((p) => p.rg));
    const report = plan.map((p, k) => {
        const conflicts = [];
        const changes = [];
        p.values.forEach((row, i) => row.forEach((v, j) => {
            const cur = current[k][i][j];
            const addr = cellAddr(p.rg, i, j);
            if (p.expected && !matches(cur, p.expected[i][j]))
                conflicts.push({ address: addr, current: preview(cur), currentHash: `sha256:${hash(cur)}` });
            if (String(cur ?? "") !== String(v ?? ""))
                changes.push({ address: addr, from: preview(cur), to: preview(v) });
        }));
        return { range: p.rg.address, status: conflicts.length ? "conflict" : changes.length ? "will-write" : "unchanged", conflicts, changes };
    });
    if (!send)
        return { dryRun: true, sheet, blocks: report, note: "Nothing written. Call again with send=true after the user approved." };
    const todo = plan.map((p, k) => ({ p, r: report[k] })).filter(({ r }) => r.status === "will-write");
    const res = await wbx.batch(todo.map(({ p }) => ({ method: "PATCH", url: wbx.rangeUrl(sheet, p.rg.address), body: { values: p.values } })));
    // Read back everything (a 504 may still have applied the write), retry once what didn't land.
    let after = await readValues(wbx, sheet, todo.map(({ p }) => p.rg));
    const landed = (k) => todo[k].p.values.every((row, i) => row.every((v, j) => normEol(after[k][i][j]) === normEol(v)));
    const retry = todo.map((_, k) => k).filter((k) => !landed(k) && res[k].status >= 500);
    if (retry.length) {
        await wbx.batch(retry.map((k) => ({ method: "PATCH", url: wbx.rangeUrl(sheet, todo[k].p.rg.address), body: { values: todo[k].p.values } })));
        after = await readValues(wbx, sheet, todo.map(({ p }) => p.rg));
    }
    todo.forEach(({ p, r }, k) => {
        const bad = [];
        p.values.forEach((row, i) => row.forEach((v, j) => {
            const got = after[k][i][j];
            if (String(got) === String(v) || normEol(got) === normEol(v))
                return;
            bad.push({ address: cellAddr(p.rg, i, j), wanted: preview(v), got: preview(got) });
        }));
        r.status = bad.length ? "written-mismatch" : "written-verified";
        if (bad.length)
            r.mismatches = bad;
        if (res[k].status >= 400 && !bad.length)
            r.note = `request returned ${res[k].status} but the read-back matches`;
        if (res[k].status >= 400 && bad.length)
            r.error = JSON.stringify(res[k].body?.error ?? res[k].body).slice(0, 300);
        r.hashes = after[k].flat().map((v) => `sha256:${hash(v)}`);
    });
    return { dryRun: false, sheet, blocks: report };
}
export function composeUpdate(cur, text, mode = "replace") {
    const s = String(cur ?? "");
    if (!s.trim())
        throw new Error("cell is empty — there is no original part to keep; use write_values");
    if (s.includes("\r\n"))
        return composeUpdate(normEol(s), text, mode);
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
export async function appendUpdate({ file, sheet, row, update, expected = {}, mode = "replace", fill = "#FFFF00", fillRange, send = false }) {
    const r = Number(row);
    if (!Number.isInteger(r) || r < 1)
        throw new Error("row must be a positive integer");
    const keys = Object.keys(update ?? {});
    const cols = keys.map((c) => c.toUpperCase());
    if (!cols.length || cols.some((c) => !/^[A-Z]{1,3}$/.test(c)))
        throw new Error('update must map column letters to text, e.g. {"H": "..."}');
    if (!["replace", "append"].includes(mode))
        throw new Error('mode must be "replace" or "append"');
    const wbx = await open(file);
    const current = (await readValues(wbx, sheet, cols.map((c) => parseRange(`${c}${r}`)))).map((v) => v[0][0]);
    const exp = Object.fromEntries(Object.entries(expected ?? {}).map(([k, v]) => [k.toUpperCase(), v]));
    const cells = cols.map((c, k) => {
        const cur = current[k];
        const address = `${c}${r}`;
        if (exp[c] !== undefined && !matches(cur, exp[c]))
            return { address, status: "conflict", current: preview(cur), currentHash: `sha256:${hash(cur)}` };
        const { head, next } = composeUpdate(cur, update[keys[k]], mode);
        return { address, status: next === String(cur) ? "unchanged" : "will-write", head, next, cur };
    });
    const fr = parseRange(fillRange || `${colName(Math.min(...cols.map(colNum)))}${r}:${colName(Math.max(...cols.map(colNum)))}${r}`);
    const shown = () => cells.map((c) => c.status === "conflict"
        ? c
        : { address: c.address, status: c.status, from: preview(c.cur, 60), updatePart: preview(c.next.slice(c.next.indexOf(SEP)), 300) });
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
    const planned = cells;
    const todo = planned.filter((c) => c.status === "will-write");
    const result = todo.length
        ? await writeValues({ file, sheet, send: true, writes: todo.map((c) => ({ range: c.address, values: c.next, expected: c.cur === null ? null : String(c.cur) })) })
        : { blocks: [] };
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
/** color: "#FFFF00" etc., or null to clear the fill. */
export async function setFill({ file, sheet, ranges, color = null, send = false }) {
    if (!Array.isArray(ranges) || !ranges.length)
        throw new Error("ranges must be a non-empty array of contiguous ranges");
    const rgs = ranges.map(parseRange);
    if (color !== null && !/^#[0-9A-Fa-f]{6}$/.test(String(color)))
        throw new Error('color must be like "#FFFF00", or null to clear');
    const want = color === null ? null : String(color).toUpperCase();
    const wbx = await open(file);
    const before = await readFills(wbx, sheet, rgs.map((r) => r.address));
    const rows = rgs.map((r, k) => ({ range: r.address, from: before[k] ?? "(none or mixed)", to: want ?? "(cleared)" }));
    if (!send)
        return { dryRun: true, sheet, ranges: rows, note: "Nothing written. Call again with send=true after the user approved." };
    await wbx.batch(rgs.map((r) => want === null
        ? { method: "POST", url: `${wbx.rangeUrl(sheet, r.address)}/format/fill/clear`, body: {} }
        : { method: "PATCH", url: `${wbx.rangeUrl(sheet, r.address)}/format/fill`, body: { color: want } }));
    const after = await readFills(wbx, sheet, rgs.map((r) => r.address));
    rows.forEach((row, k) => {
        const got = after[k] ? after[k].toUpperCase() : null;
        row.now = got ?? "(none)";
        row.status = got === want ? "verified" : "mismatch";
    });
    return { dryRun: false, sheet, ranges: rows };
}
// ---------- sheets ----------
const DRY_NOTE = "Nothing written. Call again with send=true after the user approved.";
async function sheetNames(wbx) {
    const ws = await wbx.call("GET", `${wbx.base}/worksheets?$select=name`);
    return ws.value.map((w) => w.name);
}
/** Excel's sheet-name rules; returns the trimmed name or throws. Exported for tests. */
export function checkSheetName(name, existing) {
    const n = String(name ?? "").trim();
    if (!n)
        throw new Error("Sheet name is empty");
    if (n.length > 31)
        throw new Error(`Sheet name "${n}" is longer than 31 characters`);
    if (/[\[\]:*?/\\]/.test(n) || /^'|'$/.test(n))
        throw new Error(`Sheet name "${n}" contains [ ] : * ? / \\ or starts/ends with '`);
    if (n.toLowerCase() === "history")
        throw new Error('"History" is reserved by Excel');
    const clash = existing.find((e) => e.toLowerCase() === n.toLowerCase());
    if (clash)
        throw new Error(`A sheet named "${clash}" already exists`);
    return n;
}
export async function addSheet({ file, name, position, send = false }) {
    const wbx = await open(file);
    const before = await sheetNames(wbx);
    const n = checkSheetName(name, before);
    if (position !== undefined && (!Number.isInteger(position) || position < 0 || position > before.length))
        throw new Error(`position must be 0..${before.length} (0 = first)`);
    const plan = { name: n, position: position ?? before.length, sheetsBefore: before };
    if (!send)
        return { dryRun: true, ...plan, note: DRY_NOTE };
    await wbx.call("POST", `${wbx.base}/worksheets/add`, { name: n });
    if (position !== undefined)
        await wbx.call("PATCH", wbx.sheetUrl(n), { position });
    const after = await sheetNames(wbx);
    return { dryRun: false, name: n, status: after.includes(n) ? "created-verified" : "missing-after-create", sheets: after };
}
export async function renameSheet({ file, sheet, newName, send = false }) {
    const wbx = await open(file);
    const before = await sheetNames(wbx);
    if (!before.includes(sheet))
        throw new Error(`No sheet named "${sheet}" — sheets: ${before.join(", ")}`);
    const n = checkSheetName(newName, before.filter((s) => s !== sheet));
    if (!send)
        return { dryRun: true, from: sheet, to: n, note: DRY_NOTE };
    await wbx.call("PATCH", wbx.sheetUrl(sheet), { name: n });
    const after = await sheetNames(wbx);
    return { dryRun: false, from: sheet, to: n, status: after.includes(n) && !after.includes(sheet) ? "renamed-verified" : "mismatch", sheets: after };
}
// ---------- rows ----------
/** Used range of a sheet (values only) as a Range, or null when the sheet is empty. */
async function usedRange(wbx, sheet) {
    const u = await wbx.call("GET", `${wbx.sheetUrl(sheet)}/usedRange(valuesOnly=true)?$select=address,values`);
    const addr = String(u.address).replace(/^.*!/, "");
    const rg = parseRange(addr);
    const empty = rg.rows === 1 && rg.cols === 1 && (u.values?.[0]?.[0] ?? "") === "";
    return empty ? null : rg;
}
/** Row r across the used columns (at least column A). */
const rowRange = (r, used) => parseRange(`A${r}:${colName(used ? used.c1 + used.cols - 1 : 1)}${r}`);
const rowHash = (row) => `sha256:${hash(JSON.stringify(row.map((v) => normEol(v))))}`;
export async function insertRows({ file, sheet, row, count = 1, send = false }) {
    const r = Number(row), n = Number(count);
    if (!Number.isInteger(r) || r < 1)
        throw new Error("row must be a positive integer (the new rows go above it)");
    if (!Number.isInteger(n) || n < 1 || n > 1000)
        throw new Error("count must be 1..1000");
    const wbx = await open(file);
    const used = await usedRange(wbx, sheet);
    const [cur] = await readValues(wbx, sheet, [rowRange(r, used)]);
    const plan = { sheet, insertAt: r, count: n, shiftsDown: { row: r, preview: preview(cur[0].join(" | "), 120), to: r + n } };
    if (!send)
        return { dryRun: true, ...plan, note: DRY_NOTE };
    await wbx.call("POST", `${wbx.sheetUrl(sheet)}/range(address='${r}:${r + n - 1}')/insert`, { shift: "Down" });
    // Verify: the old row r now sits at r+n, and the inserted rows are empty.
    const [moved, blank] = await readValues(wbx, sheet, [rowRange(r + n, used), parseRange(`A${r}:${colName(used ? used.c1 + used.cols - 1 : 1)}${r + n - 1}`)]);
    const ok = rowHash(moved[0]) === rowHash(cur[0]) && blank.every((row) => row.every((v) => v === "" || v === null));
    return { dryRun: false, ...plan, status: ok ? "inserted-verified" : "mismatch" };
}
/** First empty row below the used range (1 for an empty sheet). Exported for tests. */
export const nextRow = (used) => (used ? used.r1 + used.rows : 1);
export async function appendRows({ file, sheet, values, column = "A", send = false }) {
    if (!Array.isArray(values) || !values.length || !values.every(Array.isArray))
        throw new Error("values must be a non-empty 2D array (one array per row)");
    const width = Math.max(...values.map((v) => v.length));
    const grid = values.map((v) => [...v, ...Array(width - v.length).fill("")]);
    const c = String(column).toUpperCase();
    if (!/^[A-Z]{1,3}$/.test(c))
        throw new Error('column must be a column letter like "A"');
    const wbx = await open(file);
    const start = nextRow(await usedRange(wbx, sheet));
    const range = `${c}${start}:${colName(colNum(c) + width - 1)}${start + grid.length - 1}`;
    // Base check: every target cell must still be empty when written.
    const res = await writeValues({ file, sheet, send, writes: [{ range, values: grid, expected: grid.map((row) => row.map(() => "")) }] });
    return { ...res, appendedAt: range };
}
/** Font/format patches from the arguments; throws on invalid values. Exported for tests. */
export function formatPatch(a) {
    const font = {};
    const format = {};
    if (a.bold !== undefined)
        font.bold = Boolean(a.bold);
    if (a.italic !== undefined)
        font.italic = Boolean(a.italic);
    if (a.fontColor !== undefined) {
        if (!/^#[0-9A-Fa-f]{6}$/.test(String(a.fontColor)))
            throw new Error('fontColor must be like "#C00000"');
        font.color = String(a.fontColor).toUpperCase();
    }
    if (a.fontSize !== undefined) {
        if (!(Number(a.fontSize) >= 1 && Number(a.fontSize) <= 409))
            throw new Error("fontSize must be 1..409");
        font.size = Number(a.fontSize);
    }
    if (a.wrapText !== undefined)
        format.wrapText = Boolean(a.wrapText);
    if (a.horizontalAlignment !== undefined) {
        if (!["General", "Left", "Center", "Right", "Justify"].includes(a.horizontalAlignment))
            throw new Error("horizontalAlignment must be General, Left, Center, Right or Justify");
        format.horizontalAlignment = a.horizontalAlignment;
    }
    const numberFormat = a.numberFormat === undefined ? undefined : String(a.numberFormat);
    if (!Object.keys(font).length && !Object.keys(format).length && numberFormat === undefined)
        throw new Error("Nothing to change: pass bold, italic, fontColor, fontSize, wrapText, horizontalAlignment or numberFormat");
    return { font, format, numberFormat };
}
export async function setFormat({ file, sheet, ranges, send = false, ...opts }) {
    if (!Array.isArray(ranges) || !ranges.length)
        throw new Error("ranges must be a non-empty array of contiguous ranges");
    const rgs = ranges.map(parseRange);
    const { font, format, numberFormat } = formatPatch(opts);
    const wbx = await open(file);
    const readFonts = async () => (await wbx.batch(rgs.map((r) => ({ method: "GET", url: `${wbx.rangeUrl(sheet, r.address)}/format/font` })))).map((r, i) => wbx.ok(r, `read font ${rgs[i].address}`));
    const pick = (f) => Object.fromEntries(Object.keys(font).map((k) => [k, typeof f[k] === "string" ? f[k].toUpperCase() : f[k]]));
    const before = Object.keys(font).length ? await readFonts() : [];
    const change = { font, ...(Object.keys(format).length ? { format } : {}), ...(numberFormat !== undefined ? { numberFormat } : {}) };
    if (!send)
        return { dryRun: true, sheet, ranges: rgs.map((r, k) => ({ range: r.address, fontNow: before[k] ? pick(before[k]) : undefined })), change, note: DRY_NOTE };
    const reqs = [];
    for (const r of rgs) {
        if (Object.keys(font).length)
            reqs.push({ method: "PATCH", url: `${wbx.rangeUrl(sheet, r.address)}/format/font`, body: font });
        if (Object.keys(format).length)
            reqs.push({ method: "PATCH", url: `${wbx.rangeUrl(sheet, r.address)}/format`, body: format });
        if (numberFormat !== undefined)
            reqs.push({ method: "PATCH", url: wbx.rangeUrl(sheet, r.address), body: { numberFormat: Array.from({ length: r.rows }, () => Array(r.cols).fill(numberFormat)) } });
    }
    const res = await wbx.batch(reqs);
    const errors = res.filter((r) => r.status >= 400).map((r) => JSON.stringify(r.body?.error ?? r.body).slice(0, 200));
    // Read everything back: font, format (wrap/alignment) and the number format of every cell.
    const after = Object.keys(font).length ? await readFonts() : [];
    const fmtAfter = Object.keys(format).length
        ? (await wbx.batch(rgs.map((r) => ({ method: "GET", url: `${wbx.rangeUrl(sheet, r.address)}/format` })))).map((r, i) => wbx.ok(r, `read format ${rgs[i].address}`))
        : [];
    const nfAfter = numberFormat !== undefined
        ? (await wbx.batch(rgs.map((r) => ({ method: "GET", url: `${wbx.rangeUrl(sheet, r.address)}?$select=numberFormat` })))).map((r, i) => wbx.ok(r, `read numberFormat ${rgs[i].address}`).numberFormat)
        : [];
    const want = pick(font);
    return {
        dryRun: false,
        sheet,
        ranges: rgs.map((r, k) => {
            const fontOk = !after[k] || JSON.stringify(pick(after[k])) === JSON.stringify(want);
            const formatOk = !fmtAfter[k] || Object.entries(format).every(([key, v]) => fmtAfter[k][key] === v);
            const nfOk = !nfAfter[k] || nfAfter[k].every((row) => row.every((v) => v === numberFormat));
            return {
                range: r.address,
                status: errors.length ? "error" : fontOk && formatOk && nfOk ? "applied-verified" : "mismatch",
                ...(after[k] ? { fontNow: pick(after[k]) } : {}),
                ...(fmtAfter[k] ? { formatNow: Object.fromEntries(Object.keys(format).map((key) => [key, fmtAfter[k][key]])) } : {}),
                ...(nfAfter[k] && !nfOk ? { numberFormatNow: nfAfter[k] } : {}),
            };
        }),
        ...(errors.length ? { errors } : {}),
    };
}
/** Cells in `grid` (top-left at `origin`) whose text matches. Exported for tests. */
export function findInGrid(sheet, origin, grid, test, limit) {
    const hits = [];
    for (let i = 0; i < grid.length && hits.length < limit; i++)
        for (let j = 0; j < grid[i].length && hits.length < limit; j++) {
            const v = grid[i][j];
            if (v === null || v === "")
                continue;
            if (test(normEol(v)))
                hits.push({ sheet, address: cellAddr(origin, i, j), value: preview(v, 160), hash: `sha256:${hash(v)}` });
        }
    return hits;
}
export async function find({ file, text, sheet, regex = false, matchCase = false, limit = 50 }) {
    const q = String(text ?? "");
    if (!q)
        throw new Error("text is empty");
    const max = Math.min(Math.max(1, Number(limit) || 50), 500);
    let test;
    if (regex) {
        const re = new RegExp(q, matchCase ? "" : "i");
        test = (s) => re.test(s);
    }
    else {
        const needle = matchCase ? q : q.toLowerCase();
        test = (s) => (matchCase ? s : s.toLowerCase()).includes(needle);
    }
    const wbx = await open(file);
    const names = sheet ? [sheet] : await sheetNames(wbx);
    const res = await wbx.batch(names.map((s) => ({ method: "GET", url: `${wbx.sheetUrl(s)}/usedRange(valuesOnly=true)?$select=address,values` })));
    const hits = [];
    const skipped = [];
    res.forEach((r, k) => {
        if (r.status >= 400)
            return void skipped.push(names[k]);
        const origin = parseRange(String(r.body.address).replace(/^.*!/, ""));
        hits.push(...findInGrid(names[k], origin, r.body.values, test, max - hits.length));
    });
    return { text: q, hits, truncated: hits.length >= max, ...(skipped.length ? { unreadableSheets: skipped } : {}) };
}
