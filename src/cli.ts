#!/usr/bin/env node
// Entry point: `graph-excel-mcp` serves MCP over stdio; `graph-excel-mcp login` signs in from a terminal.
import { cliLogin } from "./graph.ts";
import { serve } from "./mcp.ts";

if (process.argv[2] === "login") {
  cliLogin().catch((e: Error) => {
    console.error(e.message);
    process.exit(1);
  });
} else {
  serve();
}
