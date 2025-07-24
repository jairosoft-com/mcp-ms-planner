interface DateTimeTimeZone {
    dateTime: string;
    timeZone: string;
}

interface EmailAddress {
    name?: string;
    address?: string;
}

interface Attendee {
    emailAddress: EmailAddress;
    type: 'required' | 'optional' | 'resource';
}

export interface PlannerTask {
    id?: string;
    title?: string;
    planId?: string;
    bucketId?: string;
    dueDateTime?: string | null;
    percentComplete?: number;
    value?: string; 
}

export interface GraphResponse<T> {
    value: T[];
    '@odata.nextLink'?: string;
}

// Define environment variable types
export interface Env {
    AUTH_TOKEN?: string;
}