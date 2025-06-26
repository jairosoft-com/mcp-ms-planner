import type { McpServer } from '@modelcontextprotocol/sdk/server/mcp.js';
import { createContactSchema } from '../schemas/contactSchemas.js';
import { createContact } from '../services/contactsService.js';
import type { CreateContactInput } from '../schemas/contactSchemas.js';

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
        const contactData = args as CreateContactInput;
        
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
}

/**
 * Formats a contact object into a readable string
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
