// Builds both distributions from the compiled dist/ (run `npm run build` first; `npm run pack` does both):
//   build/graph-excel-mcp-<version>.mcpb  Claude Desktop extension (chat)
//   plugin/                               Claude plugin (Cowork, Claude Code): skill + MCP server
// The token stays in ~/.graph-excel (GRAPH_EXCEL_HOME), outside both, so updates keep the sign-in.
import { execFileSync } from "node:child_process";
import fs from "node:fs";
import path from "node:path";
import { TOOLS } from "../src/mcp.ts";

const ROOT = path.join(import.meta.dirname, "..");
const pkg = JSON.parse(fs.readFileSync(path.join(ROOT, "package.json"), "utf8"));
const REPO = "https://github.com/mickey-ichi/graph-excel-mcp";
const AUTHOR = { name: "mickey-ichi", url: "https://github.com/mickey-ichi" };
const KEYWORDS = ["excel", "sharepoint", "microsoft-graph", "teams"];
const DESCRIPTION =
  "Read and edit Excel workbooks on SharePoint/Teams through Microsoft Graph, no browser: base-checked writes, " +
  "read-back verification, 「-----update-----」 test-case updates with yellow fill.";
const json = (v: unknown) => JSON.stringify(v, null, 2) + "\n";

if (!fs.existsSync(path.join(ROOT, "dist", "cli.js"))) throw new Error("dist/ is missing — run npm run build first");

// Server layout shared by both: dist/*.js + an ESM package.json (dist/ reads its version from it).
function copyServer(dir: string) {
  fs.mkdirSync(dir, { recursive: true });
  fs.cpSync(path.join(ROOT, "dist"), path.join(dir, "dist"), { recursive: true });
  for (const f of ["LICENSE", "workbooks.example.json"]) fs.copyFileSync(path.join(ROOT, f), path.join(dir, f));
  fs.writeFileSync(path.join(dir, "package.json"), json({ name: pkg.name, version: pkg.version, type: "module", license: pkg.license }));
}

// ---------- .mcpb ----------
const stage = path.join(ROOT, "build", "mcpb");
const out = path.join(ROOT, "build", `graph-excel-mcp-${pkg.version}.mcpb`);
fs.rmSync(stage, { recursive: true, force: true });
copyServer(stage);
fs.copyFileSync(path.join(ROOT, "README.md"), path.join(stage, "README.md"));
fs.writeFileSync(
  path.join(stage, "manifest.json"),
  json({
    manifest_version: "0.3",
    name: "graph-excel",
    display_name: "Graph Excel",
    version: pkg.version,
    description: DESCRIPTION,
    long_description:
      "Connects Claude to Excel workbooks on SharePoint/OneDrive through Microsoft Graph. Sign in once with a " +
      "device code (login_status gives you the code); the refresh token is kept in the data folder, not in the " +
      "extension. Every write is a dry run until approved, checks the cells' current content first and reads them " +
      "back afterwards.",
    author: AUTHOR,
    repository: { type: "git", url: REPO },
    homepage: REPO,
    support: `${REPO}/issues`,
    license: pkg.license,
    keywords: KEYWORDS,
    server: {
      type: "node",
      entry_point: "dist/cli.js",
      mcp_config: { command: "node", args: ["${__dirname}/dist/cli.js"], env: { GRAPH_EXCEL_HOME: "${user_config.home_dir}" } },
    },
    user_config: {
      home_dir: {
        type: "directory",
        title: "Data folder",
        description: "Holds token.json (sign-in) and an optional workbooks.json with your own workbook aliases.",
        required: false,
        default: "${HOME}/.graph-excel",
      },
    },
    tools: TOOLS.map(({ name, description }) => ({ name, description })),
    // No `compatibility` block: Claude Desktop runs it on its built-in Node, so users need no Node install.
  }),
);
const mcpb = (...args: string[]) => execFileSync("npx", ["--yes", "@anthropic-ai/mcpb@2.1.2", ...args], { stdio: "inherit" });
mcpb("validate", path.join(stage, "manifest.json"));
fs.rmSync(out, { force: true });
mcpb("pack", stage, out);

// ---------- plugin/ ----------
const PLUGIN = path.join(ROOT, "plugin");
fs.rmSync(PLUGIN, { recursive: true, force: true });
fs.mkdirSync(path.join(PLUGIN, ".claude-plugin"), { recursive: true });
fs.writeFileSync(
  path.join(PLUGIN, ".claude-plugin", "plugin.json"),
  json({ name: "graph-excel", displayName: "Graph Excel", version: pkg.version, description: DESCRIPTION, author: AUTHOR, homepage: REPO, repository: REPO, keywords: KEYWORDS, license: pkg.license }),
);
fs.writeFileSync(path.join(PLUGIN, ".mcp.json"), json({ mcpServers: { "graph-excel": { command: "node", args: ["${CLAUDE_PLUGIN_ROOT}/server/dist/cli.js"] } } }));
copyServer(path.join(PLUGIN, "server"));
fs.cpSync(path.join(ROOT, "skills"), path.join(PLUGIN, "skills"), { recursive: true, dereference: true });
fs.copyFileSync(path.join(ROOT, "README.md"), path.join(PLUGIN, "README.md"));

console.log(`\n${path.relative(ROOT, out)}\nplugin/`);
