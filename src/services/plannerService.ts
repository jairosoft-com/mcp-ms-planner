import { getGraphClient } from './authService.js';
import type { PlannerTask } from '../interfaces/plannerTask.js';
import type { FetchPlannerTasksInput } from '../schemas/fetchPlannerTasksSchema.js';

declare const fetch: typeof globalThis.fetch;

// Default user ID if not provided
const DEFAULT_USER_ID = process.env.USER_ID || 'me';

/**
 * Fetches planner tasks for the current user
 * @param input Query parameters for filtering tasks
 * @returns Promise with the list of tasks and pagination info
 */
export async function fetchPlannerTasks(
  input: FetchPlannerTasksInput = { user_id: 'me', top: 100, skip: 0 }
): Promise<{
  tasks: PlannerTask[];
  count: number;
  hasMore: boolean;
}> {  
  // Merge input with defaults
  const { user_id = 'me', status, top = 100, skip = 0 } = input;
  const graphClient = getGraphClient();
  
  try {
    // Build the filter string based on input parameters
    const filterParts = [];
    
    if (status) {
      // Map status to percentComplete values:
      // notStarted = 0, inProgress = 1-99, completed = 100
      if (status === 'completed') {
        filterParts.push('percentComplete eq 100');
      } else if (status === 'notStarted') {
        filterParts.push('percentComplete eq 0');
      } else if (status === 'inProgress') {
        filterParts.push('percentComplete gt 0 and percentComplete lt 100');
      }
    }
    
    const filter = filterParts.length > 0 ? filterParts.join(' and ') : undefined;
    
    // Determine the user ID to use - use environment variable if 'me' is specified
    const userId = user_id === 'me' ? (process.env.USER_ID || 'me') : user_id;
    
    // Fetch tasks from Microsoft Graph for the current user
    let request = graphClient
      .api(`/users/${userId}/planner/tasks`)
      .header('Prefer', 'odata.maxpagesize=100')
      .top(top)
      .skip(skip);
    
    if (filter) {
      request = request.filter(filter);
    }
    
    // Default user ID if not provided
    const DEFAULT_USER_ID = process.env.USER_ID || 'me';
    
    // Get the user's tasks
    const response = await request.get();
    const tasks: PlannerTask[] = response.value || [];
    
    // Check if there are more results
    const nextLink = response['@odata.nextLink'];
    const hasMore = !!nextLink;
    
    return {
      tasks,
      count: tasks.length,
      hasMore,
    };
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    console.error('Error fetching planner tasks:', error);
    throw new Error(`Failed to fetch planner tasks: ${errorMessage}`);
  }
}

/**
 * Gets detailed information about a specific task
 * @param taskId The ID of the task to get details for
 * @returns Promise with the task details
 */
export async function getTaskDetails(taskId: string): Promise<PlannerTask> {
  const graphClient = getGraphClient();
  
  try {
    const task = await graphClient
      .api(`/planner/tasks/${taskId}`)
      .select('*')
      .get();
      
    return task;
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    console.error(`Error fetching task details for task ${taskId}:`, error);
    throw new Error(`Failed to fetch task details: ${errorMessage}`);
  }
}

export default {
  fetchPlannerTasks,
  getTaskDetails,
};
