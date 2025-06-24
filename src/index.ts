// @ts-ignore - Missing type definitions for @modelcontextprotocol/sdk
import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
// @ts-ignore - Missing type definitions for @modelcontextprotocol/sdk
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import { registerPlannerTools } from './tools/plannerTools.js';
import dotenv from 'dotenv';
import path from 'path';
import { fileURLToPath } from 'url';

// Get the directory name of the current module
const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Load environment variables from .env file in the project root
const envPath = path.resolve(__dirname, '../../.env');
dotenv.config({ path: envPath });

// Check for required environment variables
const requiredEnvVars = ["AZURE_TENANT_ID", "AZURE_CLIENT_ID", "AZURE_CLIENT_SECRET"];
const missingVars = requiredEnvVars.filter((varName) => !process.env[varName]);

if (missingVars.length > 0) {
  // Error messages removed for production
  process.exit(1);
}

// Environment variables loaded

// Create server instance
const server = new McpServer({
  name: "ms-planner",
  version: "1.0.0",
});

// Register tools
registerPlannerTools(server);

// Start the server
async function main() {
  try {
    // Server startup logging removed for production
    
    const transport = new StdioServerTransport();
    await server.connect(transport);
  } catch (error) {
    // Server failed to start
    throw error;
  }
}

main().catch((error) => {
  // Fatal error in main
  process.exit(1);
});