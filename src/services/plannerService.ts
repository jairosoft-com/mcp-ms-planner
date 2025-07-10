import { authenticate, getUserIdFromToken, getGraphClient } from './authService.js';
import type { PlannerTask } from '../interfaces/plannerTask.js';
import type { FetchPlannerTasksInput } from '../schemas/fetchPlannerTasksSchema.js';
import type { CreatePlannerTaskInput } from '../schemas/createPlannerTaskSchema.js';
import type { AuthConfig } from './authService.js';

declare const fetch: typeof globalThis.fetch;

/**
 * Fetches planner tasks for the current user
 * @param input Query parameters for filtering tasks
 * @returns Promise with the list of tasks and pagination info
 */
export async function fetchPlannerTasks(
  input: FetchPlannerTasksInput & { accessToken: string } = { accessToken: '' }
): Promise<{
  tasks: PlannerTask[];
  count: number;
  hasMore: boolean;
}> {  
  const { status, accessToken } = input;
  
  // Get the user ID from the access token
  const userId = getUserIdFromToken(accessToken);
  if (!userId) {
    throw new Error('Could not determine user ID from access token');
  }
  
  const graphClient = getGraphClient(accessToken);
  
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
    
    // Build the base request using the explicit user ID
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
    console.error('Error fetching planner tasks:', error);
    throw new Error(`Failed to fetch planner tasks: ${errorMessage}`);
  }
}

/**
 * Gets detailed information about a specific task
 * @param taskId The ID of the task to get details for
 * @param authConfig Authentication configuration containing the access token
 * @returns Promise with the task details
 */
export async function getTaskDetails(
  taskId: string,
  accessToken: string
): Promise<PlannerTask> {
  try {
    // Get the user ID from the access token
    const userId = getUserIdFromToken(accessToken);
    if (!userId) {
      throw new Error('Could not determine user ID from access token');
    }
    
    const graphClient = getGraphClient(accessToken);
    
    try {
      // Get task details using the explicit user ID
      const task = await graphClient
        .api(`/users/${userId}/planner/tasks/${taskId}`)
        .select('*')
        .get();
        
      return task;
    } catch (error: unknown) {
      const errorMessage = error instanceof Error ? error.message : 'Unknown error';
      console.error(`Error fetching task details for task ${taskId}:`, error);
      throw new Error(`Failed to fetch task details: ${errorMessage}`);
    }
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    console.error(`Error fetching task details for task ${taskId}:`, error);
    throw new Error(`Failed to fetch task details: ${errorMessage}`);
  }
}

/**
 * Creates a new task in Microsoft Planner
 * @param input The task creation parameters including authConfig
 * @returns Promise with the created task
 */
export async function createPlannerTask(
  input: CreatePlannerTaskInput & { accessToken: string }
): Promise<PlannerTask> {
  const { 
    planId, 
    bucketId, 
    title = 'New Task', 
    startDateTime,
    notes,
    priority, // This is already transformed to a number by the schema
    accessToken,
    ...rest 
  } = input;
  
  // Get the user ID from the token
  const userIdFromToken = getUserIdFromToken(accessToken);
  if (!userIdFromToken) {
    throw new Error('Could not determine user ID from access token');
  }
  
  const graphClient = getGraphClient(accessToken);
  
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
    
    // Always assign the task to the current user
    taskData.assignments = {
      [userIdFromToken]: {
        '@odata.type': '#microsoft.graph.plannerAssignment',
        orderHint: ' !'
      }
    };
    
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
    console.error('Error creating planner task:', error);
    throw new Error(`Failed to create planner task: ${errorMessage}`);
  }
}

export default {
  fetchPlannerTasks,
  getTaskDetails,
  createPlannerTask,
};
