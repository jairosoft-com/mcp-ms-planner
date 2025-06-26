import { Client } from "@microsoft/microsoft-graph-client";
import { getAzureCredentials } from "../services/authService.js";
import { 
  Contact, 
  ContactListResponse, 
  ListContactsOptions 
} from "../interfaces/contact.js";

export type { Contact } from "../interfaces/contact.js";

// Initialize Microsoft Graph client with Azure AD credentials
function getAuthenticatedClient() {
  const credentials = getAzureCredentials();
  
  const client = Client.initWithMiddleware({
    authProvider: {
      getAccessToken: async () => {
        const tokenResponse = await credentials.getToken("https://graph.microsoft.com/.default");
        return tokenResponse.token;
      }
    }
  });

  return client;
}

/**
 * Create a new contact
 * @param contactData Contact data to create
 * @returns The created contact with all fields populated
 * @throws Error if the contact cannot be created
 */
export async function createContact(contactData: Omit<Contact, 'id'>): Promise<Contact> {
  try {
    const client = getAuthenticatedClient();
    const userId = process.env.USER_ID || 'me'; // 'me' refers to the current authenticated user
    
    // Ensure required fields are present
    if (!contactData.givenName && !contactData.displayName) {
      throw new Error('Either givenName or displayName is required');
    }
    
    // Prepare the contact data for Microsoft Graph API
    const contactPayload: Record<string, any> = {
      givenName: contactData.givenName,
      surname: contactData.surname,
      displayName: contactData.displayName || 
                 [contactData.givenName, contactData.surname].filter(Boolean).join(' ').trim(),
      emailAddresses: contactData.emailAddresses?.map(email => ({
        address: email.address,
        name: email.name || email.address
      })),
      businessPhones: contactData.businessPhones,
      mobilePhone: contactData.mobilePhone,
      homePhones: contactData.homePhones,
      jobTitle: contactData.jobTitle,
      companyName: contactData.companyName,
      department: contactData.department,
      officeLocation: contactData.officeLocation,
    };
    
    // Remove undefined values
    Object.keys(contactPayload).forEach(key => {
      if (contactPayload[key] === undefined) {
        delete contactPayload[key];
      }
    });
    
    // Create the contact
    const response = await client
      .api(`/users/${userId}/contacts`)
      .post(contactPayload);
      
    // Format the response to match our Contact interface
    const createdContact: Contact = {
      id: response.id,
      givenName: response.givenName,
      surname: response.surname,
      displayName: response.displayName,
      emailAddresses: response.emailAddresses?.map((email: any) => ({
        address: email.address,
        name: email.name
      })),
      businessPhones: response.businessPhones,
      mobilePhone: response.mobilePhone,
      homePhones: response.homePhones,
      jobTitle: response.jobTitle,
      companyName: response.companyName,
      department: response.department,
      officeLocation: response.officeLocation,
    };
    
    return createdContact;
  } catch (error) {
    console.error('Error creating contact:', error);
    throw new Error(`Failed to create contact: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
}

/**
 * Get a contact by ID
 * @param contactId ID of the contact to retrieve
 * @returns The requested contact
 */
export async function getContact(contactId: string): Promise<Contact> {
  try {
    const client = getAuthenticatedClient();
    const userId = process.env.USER_ID || 'me';
    
    const response = await client
      .api(`/users/${userId}/contacts/${contactId}`)
      .get();
      
    return response as Contact;
  } catch (error) {
    console.error('Error getting contact:', error);
    throw new Error(`Failed to get contact: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
}

/**
 * Update an existing contact
 * @param contactId ID of the contact to update
 * @param updateData Contact data to update
 * @returns The updated contact
 */
export async function updateContact(
  contactId: string,
  updateData: Partial<Contact>
): Promise<Contact> {
  try {
    const client = getAuthenticatedClient();
    const userId = process.env.USER_ID || 'me';
    
    const response = await client
      .api(`/users/${userId}/contacts/${contactId}`)
      .patch(updateData);
      
    return response as Contact;
  } catch (error) {
    console.error('Error updating contact:', error);
    throw new Error(`Failed to update contact: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
}

/**
 * Delete a contact
 * @param contactId ID of the contact to delete
 */
export async function deleteContact(contactId: string): Promise<void> {
  try {
    const client = getAuthenticatedClient();
    const userId = process.env.USER_ID || 'me';
    
    await client
      .api(`/users/${userId}/contacts/${contactId}`)
      .delete();
  } catch (error) {
    console.error('Error deleting contact:', error);
    throw new Error(`Failed to delete contact: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
}

/**
 * List contacts with optional filtering and pagination
 * @param options Query options (filter, select, top, skip, orderBy)
 * @returns Paginated list of contacts
 */
export async function listContacts(options: {
  filter?: string;
  select?: string[];
  top?: number;
  skip?: number;
  orderBy?: string;
} = {}): Promise<{ value: Contact[]; '@odata.nextLink'?: string }> {
  try {
    const { filter, select, top, skip, orderBy } = options;
    const client = getAuthenticatedClient();
    const userId = process.env.USER_ID || 'me';
    
    let request = client.api(`/users/${userId}/contacts`);
    
    // Apply query parameters if provided
    if (filter) request = request.filter(filter);
    if (select && select.length > 0) request = request.select(select);
    if (top) request = request.top(top);
    if (skip) request = request.skip(skip);
    if (orderBy) request = request.orderby(orderBy);
    
    const response = await request.get();
    return response as { value: Contact[]; '@odata.nextLink'?: string };
  } catch (error) {
    console.error('Error listing contacts:', error);
    throw new Error(`Failed to list contacts: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
}

/**
 * Search for contacts by name or email
 * @param query Search query string
 * @returns Matching contacts
 */
export async function searchContacts(query: string): Promise<Contact[]> {
  try {
    const client = getAuthenticatedClient();
    const userId = process.env.USER_ID || 'me';
    
    // Search for contacts where displayName, givenName, surname, or email contains the query
    const filter = `startswith(displayName,'${query}') or startswith(givenName,'${query}') or startswith(surname,'${query}') or emailAddresses/any(a:startswith(a/address,'${query}'))`;
    
    const response = await client
      .api(`/users/${userId}/contacts`)
      .filter(filter)
      .get();
      
    return (response.value || []) as Contact[];
  } catch (error) {
    console.error('Error searching contacts:', error);
    throw new Error(`Failed to search contacts: ${error instanceof Error ? error.message : 'Unknown error'}`);
  }
}
