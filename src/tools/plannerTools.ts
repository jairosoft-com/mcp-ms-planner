import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { fetchPlannerTasks } from '../services/plannerService.js';
import { fetchPlannerTasksSchema } from '../schemas/fetchPlannerTasksSchema.js';
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
}

export default {
  registerPlannerTools,
};
