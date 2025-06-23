// @ts-ignore - Missing type definitions for @modelcontextprotocol/sdk
import { McpServer } from "@modelcontextprotocol/sdk/server/mcp.js";
// @ts-ignore - Missing type definitions for @modelcontextprotocol/sdk
import { StdioServerTransport } from "@modelcontextprotocol/sdk/server/stdio.js";
import { registerPlannerTools } from "./tools/plannerTools.js";
import dotenv from "dotenv";
import path from "path";
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
  console.error(`Error: Missing required environment variables: ${missingVars.join(", ")}`);
  console.error(`Please check your .env file at: ${envPath}`);
  console.log("Current environment variables:", Object.keys(process.env).filter(k => k.startsWith('AZURE_')));
  process.exit(1);
}

console.log("Successfully loaded environment variables");

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
    console.log("Starting Microsoft Planner server...");
    console.log("Azure Tenant ID:", process.env.AZURE_TENANT_ID ? "[SET]" : "[MISSING]");
    console.log("Azure Client ID:", process.env.AZURE_CLIENT_ID ? "[SET]" : "[MISSING]");
    
    const transport = new StdioServerTransport();
    await server.connect(transport);
    console.log("✅ Microsoft Planner server started and ready to accept connections");
  } catch (error) {
    console.error("❌ Failed to start server:", error);
    throw error;
  }
}

main().catch((error) => {
  console.error("Fatal error in main():", error);
  process.exit(1);
});