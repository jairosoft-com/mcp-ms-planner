import { z } from "zod";
import { EmailAddress, PhoneNumber } from "../interfaces/contact.js";

/**
 * Schema for creating a new contact
 */
export const createContactSchema = z.object({
  givenName: z.string().min(1, "First name is required"),
  surname: z.string().optional(),
  displayName: z.string().optional(),
  emailAddresses: z.array(
    z.object({
      address: z.string().email("Invalid email address"),
      name: z.string().optional(),
    })
  ).optional(),
  businessPhones: z.array(z.string()).optional(),
  mobilePhone: z.string().optional(),
  homePhones: z.array(z.string()).optional(),
  jobTitle: z.string().optional(),
  companyName: z.string().optional(),
  department: z.string().optional(),
  officeLocation: z.string().optional(),
});

/**
 * Schema for updating an existing contact
 */
export const updateContactSchema = z.object({
  givenName: z.string().min(1, "First name is required").optional(),
  surname: z.string().optional(),
  displayName: z.string().optional(),
  emailAddresses: z.array(
    z.object({
      address: z.string().email("Invalid email address"),
      name: z.string().optional(),
    })
  ).optional(),
  businessPhones: z.array(z.string()).optional(),
  mobilePhone: z.string().optional(),
  homePhones: z.array(z.string()).optional(),
  jobTitle: z.string().optional(),
  companyName: z.string().optional(),
  department: z.string().optional(),
  officeLocation: z.string().optional(),
});

/**
 * Schema for getting a contact by ID
 */
export const getContactSchema = z.object({
  contactId: z.string().min(1, "Contact ID is required"),
});

/**
 * Schema for deleting a contact
 */
export const deleteContactSchema = z.object({
  contactId: z.string().min(1, "Contact ID is required"),
});

/**
 * Schema for listing contacts with optional filtering and pagination
 */
export const listContactsSchema = z.object({
  filter: z.string().optional().describe("OData filter expression"),
  select: z.array(z.string()).optional().describe("Array of properties to include in the response"),
  top: z.number().int().positive().max(100).optional().describe("Maximum number of items to return (max 100)"),
  skip: z.number().int().nonnegative().optional().describe("Number of items to skip"),
  orderBy: z.string().optional().describe("Property to order by"),
});

/**
 * Schema for searching contacts
 */
export const searchContactsSchema = z.object({
  query: z.string().min(1, "Search query is required"),
});

// Export types for use in the application
export type CreateContactInput = z.infer<typeof createContactSchema>;
export type UpdateContactInput = z.infer<typeof updateContactSchema>;
export type GetContactInput = z.infer<typeof getContactSchema>;
export type DeleteContactInput = z.infer<typeof deleteContactSchema>;
export type ListContactsInput = z.infer<typeof listContactsSchema>;
export type SearchContactsInput = z.infer<typeof searchContactsSchema>;
