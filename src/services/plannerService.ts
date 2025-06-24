import { getGraphClient } from './authService.js';
import type { PlannerTask } from '../interfaces/plannerTask.js';
import type { FetchPlannerTasksInput } from '../schemas/fetchPlannerTasksSchema.js';
import type { CreatePlannerTaskInput } from '../schemas/createPlannerTaskSchema.js';

declare const fetch: typeof globalThis.fetch;

// Default user ID if not provided
const DEFAULT_USER_ID = process.env.USER_ID || 'me';

/**
 * Fetches planner tasks for the current user
 * @param input Query parameters for filtering tasks
 * @returns Promise with the list of tasks and pagination info
 */
export async function fetchPlannerTasks(
  input: FetchPlannerTasksInput = { user_id: 'me' }
): Promise<{
  tasks: PlannerTask[];
  count: number;
  hasMore: boolean;
}> {  
  const { user_id = 'me', status } = input;
  const graphClient = getGraphClient();
  
  try {
    // Build the filter string based on input parameters
    let filter: string | undefined;
    
    if (status) {
      // Map status to percentComplete values:
      // notStarted = 0, inProgress = 1-99, completed = 100
      if (status === 'completed') {
        filter = 'percentComplete eq 100';
      } else if (status === 'notStarted') {
        filter = 'percentComplete eq 0';
      } else if (status === 'inProgress') {
        filter = 'percentComplete gt 0 and percentComplete lt 100';
      }
    }
    
    // Determine the user ID to use - use environment variable if 'me' is specified
    const userId = user_id === 'me' ? (process.env.USER_ID || 'me') : user_id;
    
    // Build the base request
    let request = graphClient.api(`/users/${userId}/planner/tasks`)
      .header('Prefer', 'odata.maxpagesize=100')
      .header('ConsistencyLevel', 'eventual');
    
    // Apply filter if specified
    if (filter) {
      request = request.filter(filter);
    }
    
    // Execute the request
    const response = await request.get();
    
    let tasks: PlannerTask[] = Array.isArray(response.value) ? response.value : [];
    
    // Apply client-side filtering as a fallback if needed
    if (filter && tasks.length > 0) {
      tasks = tasks.filter(task => {
        if (status === 'completed') return task.percentComplete === 100;
        if (status === 'notStarted') return task.percentComplete === 0;
        if (status === 'inProgress') return task.percentComplete > 0 && task.percentComplete < 100;
        return true;
      });
    }
    
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
    // Error fetching task details
    throw new Error(`Failed to fetch task details: ${errorMessage}`);
  }
}

/**
 * Creates a new task in Microsoft Planner
 * @param input The task creation parameters
 * @returns Promise with the created task
 */
export async function createPlannerTask(
  input: CreatePlannerTaskInput
): Promise<PlannerTask> {
  const { 
    planId, 
    bucketId, 
    title = 'New Task', 
    userId = 'me',
    startDateTime,
    notes,
    priority, // This is already transformed to a number by the schema
    ...rest 
  } = input;
  
  const graphClient = getGraphClient();
  const currentUserId = process.env.USER_ID;
  
  try {
    // Prepare the task data
    const taskData: Record<string, any> = {
      planId,
      bucketId,
      title,
      ...rest
    };
    
    // Add startDateTime if provided
    if (startDateTime) {
      taskData.startDateTime = startDateTime;
    }
    
    // If a user ID is provided, prepare the assignment
    const targetUserId = userId === 'me' ? currentUserId : userId;
    if (targetUserId) {
      taskData.assignments = {
        [targetUserId]: {
          '@odata.type': '#microsoft.graph.plannerAssignment',
          orderHint: ' !'
        }
      };
    }
    
    // Create the task using Microsoft Graph API
    const createdTask = await graphClient
      .api('/planner/tasks')
      .header('Prefer', 'return=representation')
      .post(taskData);
    
    // If notes are provided, update the task details
    if (notes) {
      try {
        // First, get the etag for the task details
        const taskDetails = await graphClient
          .api(`/planner/tasks/${createdTask.id}/details`)
          .get();
        
        // Update the task details with the notes
        await graphClient
          .api(`/planner/tasks/${createdTask.id}/details`)
          .header('Prefer', 'return=representation')
          .header('If-Match', taskDetails['@odata.etag'])
          .patch({
            description: notes
          });
        
        // Update the created task to include the notes
        createdTask.notes = notes;
      } catch (error) {
        // Error updating task details
        // Don't fail the whole operation if updating notes fails
      }
    }
    
    return createdTask;
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    throw new Error(`Failed to create planner task: ${errorMessage}`);
  }
}

export default {
  fetchPlannerTasks,
  getTaskDetails,
  createPlannerTask,
};
