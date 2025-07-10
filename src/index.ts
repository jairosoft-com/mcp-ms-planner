// @ts-ignore - Missing type definitions for @modelcontextprotocol/sdk
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
// @ts-ignore - Missing type definitions for @modelcontextprotocol/sdk
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { registerPlannerTools } from './tools/plannerTools.js';

// Create server instance
const server = new McpServer({
  name: "ms-planner",
  version: "1.0.0",
  description: "Microsoft Planner Integration"
});

// Register tools
registerPlannerTools(server);

// Start the server
async function main() {
  try {
    const transport = new StdioServerTransport();
    await server.connect(transport);
    console.error("MCP Planner Server running on stdio");
  } catch (error) {
    console.error('Failed to start server:', error);
    process.exit(1);
  }
}

// Start the server
main().catch((error) => {
  console.error('Unhandled error in main:', error);
  process.exit(1);
});