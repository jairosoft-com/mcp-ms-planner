export interface PlannerTask {
    id?: string;
    title?: string;
    planId?: string;
    bucketId?: string;
    dueDateTime?: string | null;
    percentComplete?: number;
    value?: string; 
}

export interface CreateTaskParams {
    title: string;
    planId?: string;
    bucketId?: string;
    dueDateTime?: string;
    percentComplete?: number;
    description?: string;
    priority?: number;
    assignments?: Record<string, any>;
}

export interface GraphResponse<T> {
    value: T[];
    '@odata.nextLink'?: string;
}

// Define environment variable types
export interface Env {
    AUTH_TOKEN?: string;
}