import { z } from "zod";
import { getPriorityText } from "../utils/helpers";
import { PlannerTask } from "../interface/calendarInterfaces";


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
            planId: z.string().optional().describe("Optional plan ID to filter tasks from a specific plan"),
            includeRawResponse: z.boolean().optional().describe("Include raw API response in output"),
        },
        handler: async ({ 
            startDateTime, 
            endDateTime, 
            planId,
            includeRawResponse = false 
        }: { 
            startDateTime: string; 
            endDateTime: string; 
            planId?: string;
            includeRawResponse?: boolean;
        }) => {
            try {
                if (!currentAuthToken) {
                    throw new Error("Authentication token not found. Please configure the AUTH_TOKEN environment variable in your MCP server configuration.");
                }

                let url: string;
                let tasks: any[] = [];

                if (planId) {
                    // Get tasks from a specific plan
                    url = `https://graph.microsoft.com/v1.0/planner/plans/${planId}/tasks`;
                } else {
                    // Get all tasks assigned to the user
                    url = 'https://graph.microsoft.com/v1.0/me/planner/tasks';
                }

                // Make the API request
                const response = await fetch(url, {
                    headers: {
                        'Authorization': `Bearer ${currentAuthToken}`,
                        'Content-Type': 'application/json',
                    },
                });

                if (!response.ok) {
                    const errorData = await response.json() as { error?: { message?: string } };
                    throw new Error(`Microsoft Graph API error: ${errorData?.error?.message || response.statusText}`);
                }

                const responseData = await response.json() as { value: any[] };
                tasks = responseData.value || [];

                // Filter tasks by date range if they have due dates
                const startDate = new Date(startDateTime);
                const endDate = new Date(endDateTime);
                
                const filteredTasks = tasks.filter(task => {
                    if (!task.dueDateTime) {
                        // Include tasks without due dates
                        return true;
                    }
                    const taskDueDate = new Date(task.dueDateTime);
                    return taskDueDate >= startDate && taskDueDate <= endDate;
                });

                // Check if no tasks found
                if (filteredTasks.length === 0) {
                    return {
                        content: [{
                            type: "text" as const,
                            text: planId 
                                ? `No tasks found in plan ${planId} for the specified date range.`
                                : "No tasks found for the specified date range.",
                            _meta: {}
                        }]
                    };
                }

                // Format tasks list
                const tasksList = filteredTasks.map((task, index) => {
                    const dueDate = task.dueDateTime 
                        ? new Date(task.dueDateTime).toLocaleDateString('en-US', {
                            year: 'numeric',
                            month: 'short',
                            day: 'numeric',
                            hour: '2-digit',
                            minute: '2-digit'
                        })
                        : 'No due date';
                    
                    const status = task.percentComplete === 100 
                        ? '✅ Completed' 
                        : task.percentComplete && task.percentComplete > 0 
                            ? '🔄 In Progress' 
                            : '📝 Not Started';
                    
                    const priority = task.priority ? getPriorityText(task.priority) : 'No priority';
                    
                    return [
                        `${index + 1}. ${task.title || 'Untitled Task'}`,
                        `   Status: ${status} (${task.percentComplete || 0}%)`,
                        `   Due: ${dueDate}`,
                        `   Priority: ${priority}`,
                        task.planId ? `   Plan ID: ${task.planId}` : '',
                        task.bucketId ? `   Bucket ID: ${task.bucketId}` : '',
                        task.id ? `   Task ID: ${task.id}` : '',
                        task.createdBy?.user?.displayName ? `   Created by: ${task.createdBy.user.displayName}` : '',
                        task.assignments && Object.keys(task.assignments).length > 0 
                            ? `   Assigned to: ${Object.keys(task.assignments).length} user(s)` 
                            : '   Unassigned',
                        '' // Empty line for spacing
                    ].filter(Boolean).join('\n');
                }).join('\n');

                let responseText = `📋 *Tasks (${filteredTasks.length})*\n\n${tasksList}`;

                // Add raw response if requested
                if (includeRawResponse) {
                    responseText += `\n\n🔍 *Raw API Response:*\n\`\`\`json\n${JSON.stringify(responseData, null, 2)}\n\`\`\``;
                }

                return {
                    content: [{
                        type: "text" as const,
                        text: responseText,
                        _meta: {
                            taskCount: filteredTasks.length,
                            planId: planId,
                            dateRange: `${startDateTime} to ${endDateTime}`
                        }
                    }]
                };

            } catch (error) {
                const errorMessage = error instanceof Error ? error.message : 'Unknown error occurred';
                return {
                    content: [{
                        type: "text" as const,
                        text: `Error fetching planner tasks: ${errorMessage}`,
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
