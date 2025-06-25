import { z } from 'zod';

/**
 * Schema for creating a new Planner task
 * Only includes essential fields for task creation
 */
export const createPlannerTaskSchema = z.object({
  /**
   * The ID of the user to assign the task to
   * Use "me" to assign to the current user (default)
   * Or provide a specific user ID
   */
  userId: z.string().default('me'),
  /**
   * The ID of the plan to add the task to
   * Default: A default plan ID will be used if not provided
   */
  planId: z.string().min(1, 'Plan ID is required').optional(),
  
  /**
   * The ID of the bucket to add the task to
   * Default: A default bucket ID will be used if not provided
   */
  bucketId: z.string().min(1, 'Bucket ID is required').optional(),
  
  /**
   * Title of the task (default: 'New Task')
   */
  title: z.string().default('New Task'),
  
  /**
   * The date and time when the task is due (ISO 8601 format)
   * Default: 1 week from now
   */
  dueDateTime: z.string().datetime().default(() => {
    const date = new Date();
    date.setDate(date.getDate() + 7); // 1 week from now
    return date.toISOString();
  }),
  
  /**
   * The priority of the task
   * - 'high': Highest priority (maps to 2)
   * - 'medium': Medium priority (default, maps to 5)
   * - 'low': Low priority (maps to 8)
   */
  priority: z.enum(['low', 'medium', 'high']).default('medium')
    .transform(priority => {
      const priorityMap = {
        'low': 8,
        'medium': 5,
        'high': 2
      };
      return priorityMap[priority];
    }),
  
  /**
   * The date and time when work on the task should begin (ISO 8601 format)
   * If not provided, defaults to now
   */
  startDateTime: z.string().datetime().optional(),
  
  /**
   * Additional notes or description for the task
   */
  notes: z.string().optional()
});

/**
 * Input type for creating a Planner task
 */
export type CreatePlannerTaskInput = z.infer<typeof createPlannerTaskSchema>;

