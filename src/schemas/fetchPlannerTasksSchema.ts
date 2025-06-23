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
   * - notStarted: 0%
   * - inProgress: 1-99%
   * - completed: 100%
   * Other statuses like 'deferred' and 'waitingOnOthers' are not directly supported
   * by the Graph API and will be ignored.
   */
  status: z.enum(['notStarted', 'inProgress', 'completed']).optional(),
  
  /**
   * Maximum number of tasks to return. Default is 100.
   */
  top: z.number().int().positive().max(100).default(100),
  
  /**
   * Skip the first n tasks. Useful for pagination.
   */
  skip: z.number().int().min(0).default(0)
});

/**
 * Input type for fetching planner tasks
 */
export type FetchPlannerTasksInput = z.infer<typeof fetchPlannerTasksSchema>;
