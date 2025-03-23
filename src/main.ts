#!/usr/bin/env node

// Author: Fabian G. Williams
// Purpose: Main entry point to the MCP server for searching Microsoft 365 files

import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { searchTool } from './search.js';
import { initMsal } from './auth.js';

// Create the MCP server
const server = new McpServer({
  name: 'GraphSearchDriveItems',
  version: '0.0.1',
});

// Register the MCP Tool — pass raw Zod schema (works well with MCP’s expectations)
server.tool(
  searchTool.name,
  searchTool.description,
  searchTool.inputSchema.shape,
  searchTool.handler,
);

/**
 * Bootstraps the MCP server, initializes MSAL and starts the transport
 */
async function main() {
  const { TENANT_ID, CLIENT_ID, CLIENT_SECRET } = process.env;

  if (!TENANT_ID || !CLIENT_ID || !CLIENT_SECRET) {
    throw new Error(
      'Missing required environment variables: TENANT_ID, CLIENT_ID, CLIENT_SECRET',
    );
  }

  initMsal(TENANT_ID, CLIENT_ID, CLIENT_SECRET);

  const transport = new StdioServerTransport();
  await server.connect(transport);
}

main().catch((err) => {
  console.error('Fatal startup error:', err);
  process.exit(1);
});
