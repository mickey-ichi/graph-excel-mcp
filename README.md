# graph-excel-mcp

MCP server for reading and editing Excel workbooks on SharePoint, OneDrive or Teams through Microsoft Graph, without a browser. It has no dependencies and needs Node 18 or later.

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

**Claude Desktop extension / Claude plugin (Cowork):** clone the repo and run `npm run pack`. This builds `build/graph-excel-mcp-<version>.mcpb` and the `plugin/` folder. The plugin folder contains the server and a `graph-excel` skill.

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

## Notes

- Every call runs inside a workbook session. Sessionless calls on large workbooks can time out with a 504 even though the write was applied.
- Line breaks are sent as `\n`, which Excel stores as `\r\n`. If you write `\r\n` directly, the cell ends up with a stray `_x000D_`.

## License

MIT
