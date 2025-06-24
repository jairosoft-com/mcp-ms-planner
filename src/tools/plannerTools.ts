import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { fetchPlannerTasks, createPlannerTask } from '../services/plannerService.js';
import { fetchPlannerTasksSchema } from '../schemas/fetchPlannerTasksSchema.js';
import { createPlannerTaskSchema } from '../schemas/createPlannerTaskSchema.js';
import type { PlannerTask } from '../interfaces/plannerTask.js';

/**
 * Formats a date string to a more readable format
 * @param dateString ISO date string
 * @returns Formatted date string
 */
function formatDate(dateString?: string): string {
  if (!dateString) return 'No due date';
  return new Date(dateString).toLocaleDateString('en-US', {
    year: 'numeric',
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  });
}

/**
 * Gets the status text based on percent complete
 * @param percentComplete The completion percentage (0-100)
 * @returns Status text (Not Started, In Progress, or Completed)
 */
function getStatusText(percentComplete: number): string {
  if (percentComplete === 0) return 'Not Started';
  if (percentComplete === 100) return 'Completed';
  return 'In Progress';
}

/**
 * Truncates text to a maximum length, adding an ellipsis if needed
 * @param text The text to truncate
 * @param maxLength Maximum length before truncation
 * @returns Truncated text
 */
function truncate(text: string, maxLength: number): string {
  if (!text) return '';
  return text.length <= maxLength ? text : text.substring(0, maxLength - 3) + '...';
}

/**
 * Formats a date to a short string
 * @param dateString ISO date string
 * @returns Formatted date string (MM/DD/YYYY)
 */
function formatShortDate(dateString?: string): string {
  if (!dateString) return '';
  const date = new Date(dateString);
  return date.toLocaleDateString('en-US', {
    year: 'numeric',
    month: '2-digit',
    day: '2-digit'
  });
}

/**
 * Formats a list of tasks into a tabulated markdown table
 * @param tasks List of tasks to format
 * @returns Formatted markdown table of tasks
 */
function formatTasks(tasks: PlannerTask[]): string {
  if (tasks.length === 0) {
    return 'No tasks found.';
  }

  // Table header
  let table = '| # | Title | Status | Progress | Due Date | ID |\n';
  table += '| --- | --- | --- | --- | --- | --- |\n';

  // Add each task as a row
  tasks.forEach((task, index) => {
    const status = getStatusText(task.percentComplete);
    const dueDate = task.dueDateTime ? formatShortDate(task.dueDateTime) : 'No due date';
    const progress = `${task.percentComplete}%`;
    
    // Truncate title for better table readability
    const truncatedTitle = truncate(task.title, 50);
    
    table += `| ${index + 1} `;  // Row number
    table += `| ${truncatedTitle} `;  // Title
    table += `| ${status} `;  // Status
    table += `| ${progress} `;  // Progress
    table += `| ${dueDate} `;  // Due date
    table += `| ${task.id} `;  // Task ID
    table += '|\n';  // End of row
  });

  // Add a summary at the top
  const summary = `## Planner Tasks (${tasks.length} tasks)\n\n`;
  
  return summary + table;
}

/**
 * Registers the planner tools with the MCP server
 * @param server The MCP server instance
 */
export function registerPlannerTools(server: McpServer): void {
  // Register the Get Planner Tasks tool with the server
  server.tool(
    'get-planner-tasks',
    'Fetch Microsoft Planner tasks for a user with optional filtering',
    fetchPlannerTasksSchema.shape,
    async (args: unknown) => {
      try {
        // Type assertion is safe because the schema validates the input
        const params = args as Parameters<typeof fetchPlannerTasks>[0];
        
        // Fetch tasks from Microsoft Graph
        const { tasks, count, hasMore } = await fetchPlannerTasks(params);
        
        // Format the tasks for display
        const formattedTasks = formatTasks(tasks);
        
        // Create the response
        let response = `📋 Found ${count} tasks`;
        if (hasMore) {
          response += ' (more available)';
        }
        
        if (count > 0) {
          response += `:\n\n${formattedTasks}`;
        }
        
        return {
          content: [{
            type: 'text',
            text: response,
          }],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error';
        console.error('Error in get-planner-tasks tool:', error);
        return {
          content: [{
            type: 'text',
            text: `❌ Error fetching tasks: ${errorMessage}. Please check the logs for more details.`,
          }],
        };
      }
    }
  );

  // Register the Create Planner Task tool with the server
  server.tool(
    'create-planner-task',
    'Create a new task in Microsoft Planner',
    createPlannerTaskSchema.shape,
    async (args: unknown) => {
      try {
        // Type assertion is safe because the schema validates the input
        const userParams = args as Parameters<typeof createPlannerTask>[0];
        
        // If planId or bucketId is not provided, try to get from existing tasks
        if (!userParams.planId || !userParams.bucketId) {
          try {
            // Fetch existing tasks to get planId and bucketId
            const { tasks } = await fetchPlannerTasks({ user_id: 'me' });
            
            if (tasks && tasks.length > 0) {
              // Use the first task's plan and bucket IDs
              const exampleTask = tasks[0];
              
              const { planId, bucketId } = exampleTask;

              if (!planId || !bucketId) {
                throw new Error(`Existing task is missing required IDs. Found planId: ${!!planId}, bucketId: ${!!bucketId}`);
              }
              
              userParams.planId = planId;
              userParams.bucketId = bucketId;
              
            } else {
              // If no tasks found, try environment variables as fallback
              const envPlanId = process.env.DEFAULT_PLAN_ID;
              const envBucketId = process.env.DEFAULT_BUCKET_ID;
              
              if (envPlanId && envBucketId) {
                userParams.planId = envPlanId;
                userParams.bucketId = envBucketId;
              } else {
                throw new Error(
                  'No existing tasks found to get planId and bucketId. ' +
                  'Please either:\n' +
                  '1. Create a task manually first, or\n' +
                  '2. Provide planId and bucketId in the request, or\n' +
                  '3. Set DEFAULT_PLAN_ID and DEFAULT_BUCKET_ID environment variables.'
                );
              }
            }
          } catch (fetchError) {
            const errorMessage = fetchError instanceof Error ? fetchError.message : 'Unknown error';
            throw new Error(`Failed to fetch existing tasks: ${errorMessage}`);
          }
        }
        
        // Create the task using Microsoft Graph with the provided parameters
        const createdTask = await createPlannerTask(userParams);
        
        // Format the response
        const response = `✅ Task created successfully!\n\n` +
          `**Title:** ${createdTask.title}\n` +
          `**Plan ID:** ${createdTask.planId}\n` +
          `**Bucket ID:** ${createdTask.bucketId}\n` +
          `**Status:** ${getStatusText(createdTask.percentComplete || 0)}\n` +
          `**Progress:** ${createdTask.percentComplete || 0}%\n` +
          (createdTask.dueDateTime ? `**Due Date:** ${formatDate(createdTask.dueDateTime)}\n` : '') +
          `**ID:** ${createdTask.id}`;
        
        return {
          content: [{
            type: 'text',
            text: response,
          }],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error';
        console.error('Error in create-planner-task tool:', error);
        return {
          content: [{
            type: 'text',
            text: `❌ Failed to create task: ${errorMessage}. Please check the logs for more details.`,
          }],
        };
      }
    }
  );
}

export default {
  registerPlannerTools,
};
