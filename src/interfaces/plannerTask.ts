/**
 * Represents a Microsoft Planner task
 */
export interface PlannerTask {
  /**
   * The unique identifier of the task
   */
  id: string;

  /**
   * The title of the task
   */
  title: string;

  /**
   * The date and time when the task was created
   */
  createdDateTime: string;

  /**
   * The date and time when the task is due
   */
  dueDateTime?: string;

  /**
   * The date and time when the task was last modified
   */
  lastModifiedDateTime: string;

  /**
   * The status of the task
   */
  status: 'notStarted' | 'inProgress' | 'completed' | 'deferred' | 'waitingOnOthers';

  /**
   * The percentage of task completion
   */
  percentComplete: number;

  /**
   * The ID of the plan that contains this task
   */
  planId: string;

  /**
   * The ID of the bucket (task list) that contains this task
   */
  bucketId?: string;

  /**
   * The ID of the user who created the task
   */
  createdBy?: {
    user: {
      id: string;
      displayName: string;
      userPrincipalName: string;
    };
  };

  /**
   * The user to whom the task is assigned
   */
  assignments?: Record<string, {
    assignedDateTime: string;
    orderHint: string;
    userId: string;
  }>;

  /**
   * Detailed information about the task
   */
  details?: {
    description?: string;
    previewType?: string;
    checklist?: Record<string, {
      isChecked: boolean;
      title: string;
      orderHint: string;
    }>;
  };
}

/**
 * Response format for the get-planner-tasks tool
 */
export interface GetPlannerTasksResponse {
  tasks: PlannerTask[];
  count: number;
  hasMore: boolean;
}
