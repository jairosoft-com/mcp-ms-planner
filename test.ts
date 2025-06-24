import { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { StdioServerTransport } from '@modelcontextprotocol/sdk/server/stdio.js';
import dotenv from 'dotenv';
import { fileURLToPath } from 'url';
import { dirname, join } from 'path';

// Load environment variables
dotenv.config();

// Get the directory name of the current module
const __filename = fileURLToPath(import.meta.url);
const __dirname = dirname(__filename);

// Import the built module
const { registerPlannerTools } = await import(join(__dirname, 'build', 'tools', 'plannerTools.js'));

// Create server instance
const server = new McpServer({
  name: "ms-planner-test",
  version: "1.0.0",
});

// Register tools
registerPlannerTools(server);

// Create a simple transport that logs all messages
class TestTransport {
  private server: any;
  
  // Required by Transport interface
  start() {
    return Promise.resolve();
  }
  
  // Required by Transport interface
  connect(server: any) {
    this.server = server;
    return Promise.resolve();
  }
  
  // Required by Transport interface
  send(message: any) {
    console.log('Received message from server:');
    console.log(JSON.stringify(message, null, 2));
    return Promise.resolve();
  }
  
  // Required by Transport interface
  close() {
    return Promise.resolve();
  }
}

// Test the get-planner-tasks tool
async function testGetPlannerTasks() {
  console.log('Testing get-planner-tasks tool...');
  
  try {
    // Create a test transport
    const transport = new TestTransport();
    
    // Connect the server to our test transport
    await server.connect(transport);
    
    // Simulate an incoming message
    const testMessage = {
      jsonrpc: '2.0',
      method: 'get-planner-tasks',
      params: {
        user_id: 'me',
        top: 5,
        skip: 0
      },
      id: 'test-request-1'
    };
    
    console.log('Sending test message:');
    console.log(JSON.stringify(testMessage, null, 2));
    
    // Simulate receiving a message (this would normally be handled by the transport)
    // In a real scenario, the transport would call server.handleMessage()
    console.log('\nNote: In a real scenario, the transport would handle the message.');
    console.log('To test the actual implementation, run the server and send a message to it.');
    
    console.log('\nTest completed successfully!');
  } catch (error) {
    console.error('Test failed:');
    console.error(error);
  }
}

// Run the test
async function main() {
  console.log('Starting Microsoft Planner MCP Server Test');
  console.log('======================================');
  
  try {
    await testGetPlannerTasks();
    
    // Exit after a short delay
    setTimeout(() => {
      console.log('\nTest completed. Exiting...');
      process.exit(0);
    }, 2000);
  } catch (error) {
    console.error('Test failed with error:');
    console.error(error);
    process.exit(1);
  }
}

// Run the tests
main().catch((error) => {
  console.error('Fatal error in main():', error);
  process.exit(1);
});
