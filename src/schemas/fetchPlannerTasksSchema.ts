import { z } from 'zod';

/**
 * Schema for fetching planner tasks with optional filtering parameters
 */
export const fetchPlannerTasksSchema = z.object({
  /**
   * The ID of the user whose tasks to fetch. Use 'me' for the current user.
   * If not provided, it will use the USER_ID from environment variables.
   */
  user_id: z.string().optional().default('me'),
  
  /**
   * Filter tasks by status. Status is mapped to percentComplete:
   * - 'notStarted': Tasks with 0% complete (percentComplete = 0)
   * - 'inProgress': Tasks with 1-99% complete (0 < percentComplete < 100)
   * - 'completed': Tasks with 100% complete (percentComplete = 100)
   * 
   * If not specified, returns all tasks regardless of status.
   */
  status: z.enum(['notStarted', 'inProgress', 'completed'])
    .describe("Filter tasks by status. Allowed values: 'notStarted' (0%), 'inProgress' (1-99%), 'completed' (100%)")
    .optional()
});

/**
 * Input type for fetching planner tasks
 */
export type FetchPlannerTasksInput = z.infer<typeof fetchPlannerTasksSchema>;
