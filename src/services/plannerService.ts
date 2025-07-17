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
    // First, get all tasks the user has access to
    let request = graphClient.api(`/me/planner/tasks`)
      .header('Prefer', 'odata.maxpagesize=100')
      .header('ConsistencyLevel', 'eventual');
    
    // Apply status filter if specified
    if (status) {
      let statusFilter = '';
      if (status === 'completed') {
        statusFilter = 'percentComplete eq 100';
      } else if (status === 'notStarted') {
        statusFilter = 'percentComplete eq 0';
      } else if (status === 'inProgress') {
        statusFilter = 'percentComplete gt 0 and percentComplete lt 100';
      }
      
      if (statusFilter) {
        request = request.filter(statusFilter);
      }
    }
    
    // Execute the request
    const response = await request.get();
    let tasks: PlannerTask[] = Array.isArray(response.value) ? response.value : [];
    
    // Filter tasks to only include those assigned to the current user
    tasks = tasks.filter(task => {
      // Check if the task has assignments
      if (task.assignments) {
        // The assignments object has user IDs as keys
        const assignedUserIds = Object.keys(task.assignments);
        
        // Check if the current user is in the assigned users list
        // The assignments object keys are in the format: 'user_<userId>'
        return assignedUserIds.some(assignedUserId => {
          // Extract the actual user ID from the key
          const assignedId = assignedUserId.startsWith('user_') 
            ? assignedUserId.substring(5) 
            : assignedUserId;
            
          return assignedId === userId;
        });
      }
      return false;
    });
    
    // Apply client-side status filtering as a fallback if needed
    if (status) {
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
