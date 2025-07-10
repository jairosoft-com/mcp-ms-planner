import { z } from 'zod';

/**
 * Base authentication schema that requires an access token
 */
export const authSchema = z.object({
  /**
   * Access token obtained through OAuth 2.0 authentication
   * @example "eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9..."
   */
  accessToken: z.string({
    required_error: 'Access token is required',
    invalid_type_error: 'Access token must be a string',
  })
  .min(1, 'Access token cannot be empty')
  .describe('A valid OAuth 2.0 access token obtained through user authentication')
});

/**
 * Schema for fetching planner tasks with optional filtering parameters
 */
export const fetchPlannerTasksSchema = authSchema.extend({
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
/**
 * Input type for fetching planner tasks
 */
export type FetchPlannerTasksInput = z.infer<typeof fetchPlannerTasksSchema>;

/**
 * Authentication configuration type
 */
export type AuthConfig = {
  accessToken: string;
  expiresIn?: number;
};
