# MCP MS Planner - Postman Collection

This directory contains Postman collection for testing the MCP MS Planner API.

## Collection Contents

The `mcp-ms-planner.postman_collection.json` file includes the following requests:

1. **Get All Tasks (Assigned to Me)**
   - `GET /api/planner/tasks`
   - Retrieves all tasks assigned to the authenticated user

2. **Get Tasks by Status**
   - `GET /api/planner/tasks?status={status}`
   - Filter tasks by status: `notStarted`, `inProgress`, or `completed`

3. **Create New Task**
   - `POST /api/planner/tasks`
   - Creates a new task with the specified details
   - Sample payload included in the collection

4. **Get Task Details**
   - `GET /api/planner/tasks/{taskId}`
   - Retrieves detailed information about a specific task

5. **SSE Connection**
   - `GET /events`
   - Connects to the Server-Sent Events endpoint for real-time updates

## Setup Instructions

1. **Import the Collection**
   - Open Postman
   - Click "Import" and select the `mcp-ms-planner.postman_collection.json` file

2. **Configure Environment Variables**
   - Create a new environment in Postman
   - Add a variable named `accessToken` with your Microsoft Graph API access token
   - Select this environment in the environment dropdown

3. **Update Sample Data**
   - Replace the sample `planId` and `bucketId` values in the requests with your actual Planner plan and bucket IDs
   - Update the `assignments.user_*` field with the target user's ID

## Testing with SSE

To test the Server-Sent Events endpoint:

1. Use the "SSE Connection" request
2. The connection will remain open and receive real-time updates
3. Open another tab to create/update tasks and see the events in real-time

## Notes

- All API endpoints require authentication via Bearer token
- The base URL is set to `http://localhost:3000` by default
- Update the base URL in the collection variables if your API is hosted elsewhere
