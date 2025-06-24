# Microsoft Planner MCP Server

A Node.js implementation of the Model Context Protocol (MCP) server for Microsoft Planner, built with TypeScript for type safety and better developer experience.

## Features

- Fetch Microsoft Planner tasks with filtering options
- TypeScript support with full type definitions
- MCP protocol implementation
- Easy to extend and customize
- Built with modern Node.js features

## Prerequisites

- Node.js 18.x or later
- npm 9.x or later
- TypeScript 5.0 or later
- Microsoft 365 account with access to Microsoft Planner
- Azure AD App Registration with appropriate API permissions

## Setup

1. **Clone the repository**:
   ```bash
   git clone https://github.com/yourusername/mcp-server-nodejs.git
   cd mcp-server-nodejs
   ```

2. **Install dependencies**:
   ```bash
   npm install
   ```

3. **Configure environment variables**:
   - Copy `.env.example` to `.env`
   - Update the values with your Azure AD App Registration details
   ```env
   TENANT_ID=your_tenant_id
   CLIENT_ID=your_client_id
   CLIENT_SECRET=your_client_secret
   # Optional: Default user ID to use when 'me' is specified
   # USER_ID=me@example.com
   ```

4. **Azure AD App Registration**:
   - Register a new application in the [Azure Portal](https://portal.azure.com/)
   - Add the following API permissions:
     - Microsoft Graph > Delegated permissions > Tasks.Read
     - Microsoft Graph > Delegated permissions > Tasks.ReadWrite
   - Create a client secret and note it down
   - Note your Application (client) ID and Directory (tenant) ID

## Building the Project

To compile TypeScript to JavaScript:

```bash
npm run build
```

This will compile the TypeScript files and output them to the `build` directory.

## Usage

### Running the Server

After building, you can run the server using:

```bash
node build/index.js
```

### Available Tools

#### Get Planner Tasks

Fetches Microsoft Planner tasks with optional filtering.

**Parameters:**
- `user_id` (string, optional): The ID of the user or 'me' for current user (default: 'me')
- `plan_id` (string, optional): Filter tasks by plan ID
- `task_list_id` (string, optional): Filter tasks by bucket (task list) ID
- `status` (string, optional): Filter tasks by status (notStarted, inProgress, completed, deferred, waitingOnOthers)
- `top` (number, optional): Maximum number of tasks to return (default: 25, max: 100)
- `skip` (number, optional): Number of tasks to skip (for pagination, default: 0)

**Example Request:**
```json
{
  "user_id": "me",
  "top": 5,
  "status": "notStarted"
}
```

## Testing

To run the test script:

```bash
npm test
```

Or manually run the test script:

```bash
node test.js
```

## Development

### Project Structure

- `src/` - Source files
  - `interfaces/` - TypeScript interfaces
  - `schemas/` - Zod schemas for input validation
  - `services/` - Business logic and API clients
  - `tools/` - MCP tool implementations
  - `types/` - TypeScript type declarations
  - `index.ts` - Main entry point

### Adding New Tools

1. Create a new file in `src/tools/` for your tool
2. Define the tool's schema in `src/schemas/`
3. Add any necessary interfaces in `src/interfaces/`
4. Implement the tool's functionality in the service layer
5. Register the tool in `src/index.ts`

## License

MIT

## Contributing

Contributions are welcome! Please open an issue or submit a pull request.

- Source code is in the `src` directory
- Built files are output to the `build` directory
- The project uses TypeScript for type safety
