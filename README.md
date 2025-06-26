# MCP Server for Outlook Contacts

A Node.js implementation of the Model Context Protocol (MCP) server for managing Microsoft Outlook contacts, built with TypeScript for type safety and better developer experience.

## Features

- Full CRUD operations for Outlook contacts
- Microsoft Graph API integration
- TypeScript support with comprehensive type definitions
- Environment-based configuration
- Error handling and logging
- MCP protocol implementation

## Prerequisites

- Node.js 16.x or later
- npm 7.x or later
- TypeScript 4.5 or later
- Microsoft 365 account with appropriate permissions
- Azure AD application with required API permissions:
  - Contacts.ReadWrite
  - User.Read (delegated)
  - offline_access (for refresh tokens)

## Getting Started

### 1. Clone the repository

```bash
git clone https://github.com/yourusername/mcp-server-nodejs.git
cd mcp-server-nodejs
```

### 2. Install dependencies

```bash
npm install
```

### 3. Set up Azure AD Application

1. Go to the [Azure Portal](https://portal.azure.com/)
2. Navigate to "Azure Active Directory" > "App registrations" > "New registration"
3. Enter a name for your application and select the appropriate account type
4. After registration, note down the following from the "Overview" page:
   - Application (client) ID
   - Directory (tenant) ID
5. Go to "Certificates & secrets" and create a new client secret
6. Go to "API permissions" and add the following Microsoft Graph API permissions:
   - Contacts.ReadWrite
   - User.Read (delegated)
   - offline_access
7. Grant admin consent for the permissions

### 4. Configure Environment Variables

Copy the `.env.example` file to `.env` and update the values:

```bash
cp .env.example .env
```

Edit the `.env` file with your Azure AD application details:

```
AZURE_TENANT_ID=your_tenant_id_here
AZURE_CLIENT_ID=your_client_id_here
AZURE_CLIENT_SECRET=your_client_secret_here
USER_ID=me  # or specific user email/ID
```

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

The server provides the following MCP tools for managing contacts:

1. **createContact** - Create a new contact
   - Parameters: Contact details (firstName, lastName, email, etc.)

2. **getContact** - Get a contact by ID
   - Parameters: contactId

3. **updateContact** - Update an existing contact
   - Parameters: contactId, updated contact details

4. **deleteContact** - Delete a contact
   - Parameters: contactId

5. **listContacts** - List contacts with optional filtering and pagination
   - Parameters: filter, select, top, skip, orderBy

6. **searchContacts** - Search contacts by name or email
   - Parameters: query

## Development

### Project Structure

```
src/
  ├── interfaces/    # TypeScript interfaces
  ├── schemas/      # Zod validation schemas
  ├── services/     # Business logic and API clients
  ├── tools/        # MCP tool implementations
  └── index.ts      # Server entry point
```

### Environment Variables

See `.env.example` for all available environment variables.

### Testing

To run tests:

```bash
npm test
```

## Troubleshooting

### Common Issues

1. **Authentication Errors**
   - Ensure your Azure AD application has the correct permissions
   - Verify that the client secret hasn't expired
   - Check that the redirect URI is correctly configured

2. **Missing Environment Variables**
   - Ensure all required environment variables are set in your `.env` file
   - The server will fail to start if required variables are missing

3. **API Rate Limiting**
   - The Microsoft Graph API has rate limits
   - Implement proper error handling and retry logic in your client

## License

MIT

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Acknowledgments

- [Microsoft Graph API](https://learn.microsoft.com/en-us/graph/overview)
- [Model Context Protocol](https://github.com/modelcontextprotocol)
- [TypeScript](https://www.typescriptlang.org/)

- Source code is in the `src` directory
- Built files are output to the `build` directory
- The project uses TypeScript for type safety
