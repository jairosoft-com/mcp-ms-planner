import { z } from "zod";
import { GraphResponse, PlannerTask } from "../interface/calendarInterfaces";

// Shared authentication token
let currentAuthToken: string | undefined;

// Function to set the authentication token
export function setAuthToken(token: string | undefined) {
    currentAuthToken = token;
}

export function fetchTasks() {
    return {
        name: "fetchTasks",
        schema: {
            startDateTime: z.string().describe("Start date and time in ISO 8601 format (e.g., 2023-01-01T00:00:00)"),
            endDateTime: z.string().describe("End date and time in ISO 8601 format (e.g., 2023-01-31T23:59:59)"),
        },
        handler: async ({ startDateTime, endDateTime }: { startDateTime: string; endDateTime: string }) => {
            try {
                if (!currentAuthToken) {
                    throw new Error("Authentication token not found. Please configure the AUTH_TOKEN environment variable in your MCP server configuration.");
                }

                // Format the API URL with query parameters
                const url = new URL('https://graph.microsoft.com/v1.0/me/calendarView');
                url.searchParams.append('startDateTime', startDateTime);
                url.searchParams.append('endDateTime', endDateTime);
                url.searchParams.append('$select', 'subject,start,end,organizer,location');
                url.searchParams.append('$orderby', 'start/dateTime');

                // Make the API request
                const response = await fetch(url.toString(), {
                    headers: {
                        'Authorization': `Bearer ${currentAuthToken}`,
                        'Content-Type': 'application/json',
                    },
                });

                if (!response.ok) {
                    const errorData = await response.json() as { error?: { message?: string } };
                    throw new Error(`Microsoft Graph API error: ${errorData?.error?.message || response.statusText}`);
                }

                const responseData = await response.json() as PlannerTask;

                const data = responseData.value || responseData;

                // Check if data is an array and handle empty results
                if (!Array.isArray(data)) {
                    return {
                        content: [{
                            type: "text" as const,
                            text: "Error: Unexpected response format from planner API.",
                            _meta: {}
                        }]
                    };
                }
                
                if (data.length === 0) {
                    return {
                        content: [{
                            type: "text" as const,
                            text: "No tasks found in the specified plan.",
                            _meta: {}
                        }]
                    };
                }
                
                const tasksList = data.map((task, index) => {
                    const dueDate = task.dueDateTime 
                        ? new Date(task.dueDateTime).toLocaleDateString() 
                        : 'No due date';
                    const status = task.percentComplete === 100 
                        ? '✅ Completed' 
                        : task.percentComplete && task.percentComplete > 0 
                            ? '🔄 In Progress' 
                            : '📝 Not Started';
                    
                    return [
                        `${index + 1}. ${task.title || 'Untitled Task'}`,
                        `   Status: ${status} (${task.percentComplete || 0}%)`,
                        `   Due: ${dueDate}`,
                        task.planId ? `   Plan ID: ${task.planId}` : '',
                        task.bucketId ? `   Bucket ID: ${task.bucketId}` : '',
                        '' // Empty line for spacing
                    ].filter(Boolean).join('\n');
                }).join('\n');
                
                return {
                    content: [{
                        type: "text" as const,
                        text: `📋 *Tasks (${data.length})*\n\n${tasksList}`,
                        _meta: {}
                    }]
                };

            } catch (error) {
                const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
                return {
                    content: [{
                        type: "text" as const,
                        text: `Error fetching calendar events: ${errorMessage}`,
                        _meta: {}
                    }],
                    isError: true
                };
            }
        }
    };
}

export function createTask() {
    return {
        name: "createTask",
        schema: {
            title: z.string().min(1, "Task title is required"),
            planId: z.string().optional().describe("ID of the plan to add the task to"),
            bucketId: z.string().optional().describe("ID of the bucket to add the task to"),
            dueDateTime: z.string().datetime().optional().describe("Due date and time in ISO 8601 format"),
            percentComplete: z.number().min(0).max(100).optional().describe("Percentage complete (0-100)"),
            description: z.string().optional().describe("Task description"),
            priority: z.number().min(0).max(10).optional().describe("Task priority (0=no priority, 10=highest)"),
            assignments: z.record(z.object({
                '@odata.type': z.literal("#microsoft.graph.plannerAssignment"),
                orderHint: z.string().optional()
            })).optional().describe("User assignments for the task")
        },
        handler: async ({
            title,
            planId,
            bucketId,
            dueDateTime,
            percentComplete,
            description,
            priority,
            assignments
        }: {
            title: string;
            planId?: string;
            bucketId?: string;
            dueDateTime?: string;
            percentComplete?: number;
            description?: string;
            priority?: number;
            assignments?: Record<string, any>;
        }) => {
            try {
                if (!currentAuthToken) {
                    throw new Error("Authentication token not found. Please configure the AUTH_TOKEN environment variable in your MCP server configuration.");
                }

                // If planId or bucketId is not provided, try to get from existing tasks
                let targetPlanId = planId;
                let targetBucketId = bucketId;

                if (!targetPlanId || !targetBucketId) {
                    // First, get the user's plans
                    const plansResponse = await fetch('https://graph.microsoft.com/v1.0/me/planner/plans', {
                        headers: {
                            'Authorization': `Bearer ${currentAuthToken}`,
                            'Content-Type': 'application/json',
                        },
                    });

                    if (!plansResponse.ok) {
                        const errorData = await plansResponse.json() as { error?: { message?: string } };
                        throw new Error(`Failed to fetch plans: ${errorData?.error?.message || plansResponse.statusText}`);
                    }

                    const plans = await plansResponse.json() as { value: Array<{ id: string, title: string }> };
                    if (plans.value.length === 0) {
                        throw new Error("No plans found for this user. Please create a plan first.");
                    }

                    // Use the first plan by default if not specified
                    targetPlanId = targetPlanId || plans.value[0].id;

                    // Get buckets for the selected plan
                    const bucketsResponse = await fetch(`https://graph.microsoft.com/v1.0/planner/plans/${targetPlanId}/buckets`, {
                        headers: {
                            'Authorization': `Bearer ${currentAuthToken}`,
                            'Content-Type': 'application/json',
                        },
                    });

                    if (!bucketsResponse.ok) {
                        const errorData = await bucketsResponse.json() as { error?: { message?: string } };
                        throw new Error(`Failed to fetch buckets: ${errorData?.error?.message || bucketsResponse.statusText}`);
                    }

                    const buckets = await bucketsResponse.json() as { value: Array<{ id: string, name: string }> };
                    if (buckets.value.length === 0) {
                        throw new Error("No buckets found in the selected plan. Please create a bucket first.");
                    }

                    // Use the first bucket by default if not specified
                    targetBucketId = targetBucketId || buckets.value[0].id;
                }

                // Create the task payload
                const taskPayload: any = {
                    planId: targetPlanId,
                    bucketId: targetBucketId,
                    title,
                };

                // Add optional fields if provided
                if (dueDateTime) taskPayload.dueDateTime = dueDateTime;
                if (percentComplete !== undefined) taskPayload.percentComplete = percentComplete;
                if (description) taskPayload.description = description;
                if (priority !== undefined) taskPayload.priority = priority;
                if (assignments) taskPayload.assignments = assignments;

                // First, create the task
                const createResponse = await fetch('https://graph.microsoft.com/v1.0/planner/tasks', {
                    method: 'POST',
                    headers: {
                        'Authorization': `Bearer ${currentAuthToken}`,
                        'Content-Type': 'application/json',
                    },
                    body: JSON.stringify(taskPayload)
                });

                if (!createResponse.ok) {
                    const errorData = await createResponse.json() as { error?: { message?: string } };
                    throw new Error(`Failed to create task: ${errorData?.error?.message || createResponse.statusText}`);
                }

                const createdTask = await createResponse.json() as PlannerTask;

                // Helper function to get task status text
                function getTaskStatus(percentComplete?: number): string {
                    if (percentComplete === undefined || percentComplete === null) {
                        return 'Not Started (0%)';
                    }
                    const status = percentComplete === 100 
                        ? 'Completed' 
                        : percentComplete > 0 
                            ? 'In Progress' 
                            : 'Not Started';
                    return `${status} (${percentComplete}%)`;
                }

                // Format the response
                const responseText = [
                    `✅ Task created successfully!`,
                    `Title: ${createdTask?.title || 'No title'}`,
                    `ID: ${createdTask?.id || 'No ID'}`,
                    `Plan ID: ${createdTask?.planId || 'No plan ID'}`,
                    `Bucket ID: ${createdTask?.bucketId || 'No bucket ID'}`,
                    createdTask?.dueDateTime 
                        ? `Due: ${new Date(createdTask.dueDateTime).toLocaleString()}` 
                        : 'No due date',
                    `Status: ${getTaskStatus(createdTask?.percentComplete)}`
                ].join('\n');

                return {
                    content: [{
                        type: "text" as const,
                        text: responseText,
                        _meta: {}
                    }]
                };

            } catch (error) {
                const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
                return {
                    content: [{
                        type: "text" as const,
                        text: `Error creating planner task: ${errorMessage}`,
                        _meta: {}
                    }],
                    isError: true
                };
            }
        }
    };
}
