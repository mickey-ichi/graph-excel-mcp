#!/usr/bin/env node
// Builds both distributions from the same zero-dependency server files:
//   build/graph-excel-<version>.mcpb  Claude Desktop extension (chat)
//   plugin/                           Claude plugin (Cowork, Claude Code): skill + MCP server
// The token stays in ~/.graph-excel (GRAPH_EXCEL_HOME), outside both, so updates keep the sign-in.
const { execFileSync } = require("child_process");
const fs = require("fs");
const path = require("path");

process.env.GRAPH_EXCEL_NO_SERVE = "1";
const ROOT = path.join(__dirname, "..");
const { TOOLS } = require(path.join(ROOT, "mcp.js"));
const pkg = require(path.join(ROOT, "package.json"));
const SERVER_FILES = ["graph.js", "excel.js", "mcp.js", "workbooks.example.json", "LICENSE"];
const DESCRIPTION =
  "Read and edit Excel workbooks on SharePoint/Teams through Microsoft Graph, no browser: base-checked writes, " +
  "read-back verification, 「-----update-----」 test-case updates with yellow fill.";
const json = (v) => JSON.stringify(v, null, 2) + "\n";

function copyServer(dir) {
  fs.mkdirSync(dir, { recursive: true });
  for (const f of SERVER_FILES) fs.copyFileSync(path.join(ROOT, f), path.join(dir, f));
  fs.writeFileSync(path.join(dir, "package.json"), json({ name: pkg.name, version: pkg.version, license: pkg.license }));
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
    author: { name: "mickey-ichi", url: "https://github.com/mickey-ichi" },
    repository: { type: "git", url: "https://github.com/mickey-ichi/graph-excel-mcp" },
    homepage: "https://github.com/mickey-ichi/graph-excel-mcp",
    support: "https://github.com/mickey-ichi/graph-excel-mcp/issues",
    license: "MIT",
    keywords: ["excel", "sharepoint", "microsoft-graph", "teams"],
    server: {
      type: "node",
      entry_point: "mcp.js",
      mcp_config: { command: "node", args: ["${__dirname}/mcp.js"], env: { GRAPH_EXCEL_HOME: "${user_config.home_dir}" } },
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
const mcpb = (...args) => execFileSync("npx", ["--yes", "@anthropic-ai/mcpb@2.1.2", ...args], { stdio: "inherit" });
mcpb("validate", path.join(stage, "manifest.json"));
fs.rmSync(out, { force: true });
mcpb("pack", stage, out);

// ---------- plugin/ ----------
const PLUGIN = path.join(ROOT, "plugin");
fs.rmSync(PLUGIN, { recursive: true, force: true });
fs.mkdirSync(path.join(PLUGIN, ".claude-plugin"), { recursive: true });
fs.writeFileSync(
  path.join(PLUGIN, ".claude-plugin", "plugin.json"),
  json({ name: "graph-excel", displayName: "Graph Excel", version: pkg.version, description: DESCRIPTION, author: { name: "mickey-ichi", url: "https://github.com/mickey-ichi" }, homepage: "https://github.com/mickey-ichi/graph-excel-mcp", repository: "https://github.com/mickey-ichi/graph-excel-mcp", keywords: ["excel", "sharepoint", "microsoft-graph"], license: "MIT" }),
);
fs.writeFileSync(path.join(PLUGIN, ".mcp.json"), json({ mcpServers: { "graph-excel": { command: "node", args: ["${CLAUDE_PLUGIN_ROOT}/server/mcp.js"] } } }));
copyServer(path.join(PLUGIN, "server"));
fs.cpSync(path.join(ROOT, "skills"), path.join(PLUGIN, "skills"), { recursive: true, dereference: true });
fs.copyFileSync(path.join(ROOT, "README.md"), path.join(PLUGIN, "README.md"));

console.log(`\n${path.relative(ROOT, out)}\nplugin/`);
