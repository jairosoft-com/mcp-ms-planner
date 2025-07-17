// @ts-ignore - Missing type definitions for @modelcontextprotocol/sdk
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
// @ts-ignore - Missing type definitions for @modelcontextprotocol/sdk
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { registerPlannerTools } from './tools/plannerTools.js';
import { SseServer } from './server/sseServer.js';

// Create MCP server instance
const mcpServer = new McpServer({
  name: "ms-planner",
  version: "1.0.0",
  description: "Microsoft Planner Integration"
});

// Register tools
registerPlannerTools(mcpServer);

// Create SSE server instance
const sseServer = new SseServer(3002);

// Handle process termination
const shutdown = async () => {
  console.log('Shutting down servers...');
  
  try {
    await sseServer.stop();
    console.log('SSE Server stopped');
  } catch (error) {
    console.error('Error stopping SSE server:', error);
  }
  
  process.exit(0);
};

// Handle process termination signals
process.on('SIGINT', shutdown);
process.on('SIGTERM', shutdown);

// Start the servers
async function main() {
  try {
    // Start MCP server
    const transport = new StdioServerTransport();
    await mcpServer.connect(transport);
    console.log("MCP Planner Server running on stdio");
    
    // Start SSE server
    await sseServer.start();
    console.log('SSE Server started');
    
  } catch (error) {
    console.error('Failed to start servers:', error);
    await shutdown();
    process.exit(1);
  }
}

// Start the servers
main().catch((error) => {
  console.error('Unhandled error in main:', error);
  process.exit(1);
});