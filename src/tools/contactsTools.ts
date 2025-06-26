import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { createContactSchema, listContactsSchema } from '../schemas/contactSchemas.js';
import { createContact, listContacts } from '../services/contactsService.js';
import type { Contact, ListContactsOptions } from '../interfaces/contact.js';

/**
 * Registers all contact-related tools with the MCP server
 * @param server - The MCP server instance
 */
export function registerContactsTools(server: McpServer): void {
  // Register the Create Contact tool with the server
  server.tool(
    'create-contact',
    'Create a new contact in Microsoft Outlook',
    createContactSchema.shape,
    async (args: unknown) => {
      try {
        // Type assertion is safe because the schema validates the input
        const contactData = args as Omit<Contact, 'id'>;
        
        // Call the contacts service to create the contact
        const createdContact = await createContact(contactData);
        
        // Format the response
        const response = formatContactResponse(createdContact);
        
        return {
          content: [{
            type: 'text',
            text: `✅ Contact created successfully!\n\n${response}`,
          }],
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error';
        
        return {
          content: [{
            type: 'text',
            text: `❌ Failed to create contact: ${errorMessage}`,
          }],
        };
      }
    }
  );

  // Register the List Contacts tool with the server
  server.tool(
    'list-contacts',
    'List all contacts in Microsoft Outlook',
    listContactsSchema.shape,
    async (args: unknown) => {
      try {
        // Type assertion is safe because the schema validates the input
        const options = args as ListContactsOptions;
        
        // Call the contacts service to list contacts
        const { value: contacts, '@odata.nextLink': nextLink } = await listContacts({
          filter: options.filter,
          select: options.select,
          top: options.top,
          skip: options.skip,
          orderBy: options.orderBy,
        });
        
        // Format the response as a table for better LLM consumption
        if (contacts.length === 0) {
          return {
            content: [{
              type: 'text',
              text: 'No contacts found.'
            }]
          };
        }

        // Create a tabular representation of contacts
        const headers = ['ID', 'Name', 'Email', 'Mobile', 'Company', 'Title'];
        const rows = contacts.map(contact => {
          const primaryEmail = contact.emailAddresses?.[0]?.address || '';
          return [
            contact.id || '',
            [contact.givenName, contact.surname].filter(Boolean).join(' ') || contact.displayName || 'Unnamed Contact',
            primaryEmail,
            contact.mobilePhone || '',
            contact.companyName || '',
            contact.jobTitle || ''
          ];
        });

        // Format as a markdown table
        const table = formatAsMarkdownTable(headers, rows);
        
        let response = `### Contacts (${contacts.length} found)\n\n${table}`;
        
        if (nextLink) {
          response += '\n\n*Note: There are more contacts available. Use the nextLink to fetch more.*';
        }
        
        return {
          content: [{
            type: 'text',
            text: response,
          }],
          metadata: {
            nextLink,
            totalContacts: contacts.length,
            hasMore: !!nextLink
          }
        };
      } catch (error: unknown) {
        const errorMessage = error instanceof Error ? error.message : 'Unknown error';
        
        return {
          content: [{
            type: 'text',
            text: `❌ Failed to list contacts: ${errorMessage}`,
          }],
        };
      }
    }
  );
}

/**
 * Formats a contact object into a readable string
 */
/**
 * Formats data as a markdown table
 */
function formatAsMarkdownTable(headers: string[], rows: string[][]): string {
  // Create the header row
  const headerRow = `| ${headers.join(' | ')} |`;
  
  // Create the separator row
  const separator = `| ${headers.map(() => '---').join(' | ')} |`;
  
  // Format each data row
  const dataRows = rows.map(row => 
    `| ${row.map(cell => cell.replace(/\n/g, ' ').trim()).join(' | ')} |`
  );
  
  // Combine everything
  return [headerRow, separator, ...dataRows].join('\n');
}

/**
 * Formats a single contact's details into a structured object
 * This is kept for backward compatibility
 */
function formatContactResponse(contact: {
  id?: string | null;
  givenName?: string | null;
  surname?: string | null;
  displayName?: string | null;
  emailAddresses?: Array<{ address?: string | null; name?: string | null; }> | null;
  mobilePhone?: string | null;
  jobTitle?: string | null;
  companyName?: string | null;
}): string {
  const parts: string[] = [];
  
  // Add name information
  const nameParts = [];
  if (contact.givenName) nameParts.push(contact.givenName);
  if (contact.surname) nameParts.push(contact.surname);
  const fullName = nameParts.length > 0 ? nameParts.join(' ') : contact.displayName || 'Unnamed Contact';
  
  parts.push(`**Name:** ${fullName}`);
  
  // Add job title and company if available
  if (contact.jobTitle) parts.push(`**Title:** ${contact.jobTitle}`);
  if (contact.companyName) parts.push(`**Company:** ${contact.companyName}`);
  
  // Add email addresses
  if (contact.emailAddresses && contact.emailAddresses.length > 0) {
    const emails = contact.emailAddresses
      .filter(email => email?.address)
      .map(email => `- ${email.name ? `${email.name} <${email.address}>` : email.address}`)
      .join('\n');
    
    if (emails) {
      parts.push('**Emails:**');
      parts.push(emails);
    }
  }
  
  // Add mobile phone if available
  if (contact.mobilePhone) {
    parts.push(`**Mobile:** ${contact.mobilePhone}`);
  }
  
  // Add contact ID at the end
  parts.push(`\n**Contact ID:** ${contact.id}`);
  
  return parts.join('\n');
}

export default {
  registerContactsTools,
};
