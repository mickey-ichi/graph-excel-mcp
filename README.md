# graph-excel-mcp

MCP server for reading and editing Excel workbooks on SharePoint, OneDrive or Teams through Microsoft Graph, without a browser. It is written in TypeScript, has no runtime dependencies, and runs on Node 18 or later.

It is built for workbooks that several people edit at once:

- Every write starts as a dry run.
- Before writing, it checks the cells' current content against the version you read, so it won't overwrite someone else's edit.
- After writing, it reads the cells back to confirm the change.

## Install

**Claude Code / any MCP host (stdio):**

```bash
claude mcp add --scope user graph-excel -- npx -y graph-excel-mcp
```

```json
{ "mcpServers": { "graph-excel": { "command": "npx", "args": ["-y", "graph-excel-mcp"] } } }
```

**Claude Desktop extension:** download `graph-excel-mcp-<version>.mcpb` from the [latest release](https://github.com/mickey-ichi/graph-excel-mcp/releases/latest) and open it. One install covers both Chat and Cowork in Claude Desktop.

**Claude plugin (Cowork, Claude Code):** install the `plugin/` folder of this repo. It contains the server and a `graph-excel` skill.

## Sign in

Call the `login_status` tool, or run `npx graph-excel-mcp login`, then enter the code it shows at https://login.microsoft.com/device. The refresh token is saved in `~/.graph-excel/token.json` with mode 600. Set `GRAPH_EXCEL_HOME` to keep it somewhere else.

Sign-in uses the Microsoft Graph PowerShell public client with the delegated scope `Files.ReadWrite.All`, so you don't need to register an app. If your tenant blocks user consent, an admin has to approve that client once.

## Workbook aliases

Instead of pasting the SharePoint URL every time, you can add aliases to `~/.graph-excel/workbooks.json` (see `workbooks.example.json`). The `resolve` tool returns the `driveId` and `itemId` you need for an alias.

## Tools

| Tool | What it does |
|---|---|
| `login_status` | Checks the token, or starts a device code login |
| `resolve` | Returns drive/item ids, last modified, and the sheets with their used ranges |
| `read_range` | Returns values plus a `sha256:` hash per cell, and fills on request |
| `write_values` | Writes blocks of values with a base check (`expected`) and read-back verification |
| `append_update` | Adds or replaces a 「-----update-----」 section under a cell's existing text, keeping the original part byte-identical, and can fill the row with a color |
| `set_fill` | Sets or clears the fill color of one or more ranges in a single batch |
| `find` | Searches cell text across all sheets (or one), substring or regex, and returns addresses with hashes |
| `add_sheet` | Adds an empty sheet (Excel name rules, no duplicates), optionally at a position |
| `rename_sheet` | Renames a sheet |
| `insert_rows` | Inserts empty rows above a row and verifies the shifted row moved intact |
| `append_rows` | Appends rows below the last used row; the target cells must still be empty |
| `set_format` | Bold, italic, font color and size, wrap, alignment and number format for one or more ranges |

Every tool that changes the workbook is a dry run until `send: true`. Nothing deletes sheets, rows or cells.

## Notes

- Every call runs inside a workbook session. Sessionless calls on large workbooks can time out with a 504 even though the write was applied.
- Line breaks are sent as `\n`, which Excel stores as `\r\n`. If you write `\r\n` directly, the cell ends up with a stray `_x000D_`.

## Development

Development needs Node 22.18 or later, which runs the `.ts` sources directly.

| Command | What it does |
|---|---|
| `npm run mcp` / `npm run login` | Runs the server or the terminal sign-in from `src/` |
| `npm run typecheck` | Runs `tsc` in strict mode with no emit |
| `npm test` | Runs offline tests: range parsing, update composition, and the MCP handshake on `src/` and `dist/` |
| `npm run build` | Compiles `src/` to `dist/`, which is what npm, the `.mcpb` and the plugin run |
| `npm run pack` | Builds, then writes `build/*.mcpb` (gitignored) and `plugin/` (committed) |

Layout:

| Path | Contents |
|---|---|
| `src/graph.ts` | Auth and the Graph client |
| `src/excel.ts` | Workbook operations |
| `src/mcp.ts` | Tools and the JSON-RPC server |
| `src/cli.ts` | Entry point |

## Release

Bump `version` in `package.json`, run `npm run pack`, commit, then push a matching tag:

```bash
git tag v0.3.1 && git push origin v0.3.1
```

The `Publish` workflow runs the tests, publishes to npm through trusted publishing (no token or OTP; provenance attached) and creates a GitHub release with the `.mcpb` attached.

## Privacy

graph-excel-mcp has no server of its own and collects nothing. It signs in to Microsoft with a device code and keeps
the refresh token only in `~/.graph-excel/token.json` (mode 600) on your computer. It talks to Microsoft only —
`login.microsoftonline.com` to sign in and `graph.microsoft.com` to read and write the workbooks you point it at.
Cell contents go to Claude only when you ask for them, and are then handled under your Claude plan's terms. No
telemetry, no third-party services.

## Support

If graph-excel-mcp saves you time, you can support its development through
[GitHub Sponsors](https://github.com/sponsors/mickey-ichi). Bug reports and pull requests are just as welcome.

## License

MIT
