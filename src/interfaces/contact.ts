/**
 * Represents an email address for a contact
 */
export interface EmailAddress {
  /** The email address */
  address: string;
  
  /** The display name for the email address */
  name?: string;
}

/**
 * Represents a phone number for a contact
 */
export interface PhoneNumber {
  /** The phone number */
  number: string;
  
  /** The type of phone number */
  type: 'home' | 'business' | 'mobile' | 'other';
}

/**
 * Represents a contact in the system
 */
export interface Contact {
  /** The unique identifier for the contact */
  id?: string;
  
  /** The contact's first name */
  givenName?: string;
  
  /** The contact's last name */
  surname?: string;
  
  /** The contact's full name */
  displayName?: string;
  
  /** Array of email addresses */
  emailAddresses?: EmailAddress[];
  
  /** Array of business phone numbers */
  businessPhones?: string[];
  
  /** Mobile phone number */
  mobilePhone?: string;
  
  /** Array of home phone numbers */
  homePhones?: string[];
  
  /** Job title */
  jobTitle?: string;
  
  /** Company name */
  companyName?: string;
  
  /** Department */
  department?: string;
  
  /** Office location */
  officeLocation?: string;
  
  /** Date and time when the contact was created */
  createdDateTime?: string;
  
  /** Date and time when the contact was last modified */
  lastModifiedDateTime?: string;
}

/**
 * Represents a paginated list of contacts
 */
export interface ContactListResponse {
  /** Array of contacts */
  value: Contact[];
  
  /** URL to the next page of results, if available */
  '@odata.nextLink'?: string;
}

/**
 * Options for listing contacts
 */
export interface ListContactsOptions {
  /** OData filter expression */
  filter?: string;
  
  /** Array of properties to include in the response */
  select?: string[];
  
  /** Maximum number of items to return */
  top?: number;
  
  /** Number of items to skip */
  skip?: number;
  
  /** Property to order by */
  orderBy?: string;
}
