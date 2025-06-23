import { Client } from '@microsoft/microsoft-graph-client';
import { TokenCredentialAuthenticationProvider } from '@microsoft/microsoft-graph-client/authProviders/azureTokenCredentials/index.js';
import { ClientSecretCredential } from '@azure/identity';
import 'isomorphic-fetch';
import dotenv from 'dotenv';

// Load environment variables
dotenv.config();

// Type declaration for the fetch API
declare const fetch: typeof globalThis.fetch;

// Environment variables
const tenantId = process.env.AZURE_TENANT_ID;
const clientId = process.env.AZURE_CLIENT_ID;
const clientSecret = process.env.AZURE_CLIENT_SECRET;

/**
 * Validates if all required Azure AD environment variables are set
 * @throws Error if any required environment variable is missing
 */
function validateAuthConfig(): void {
  const requiredVars = ['AZURE_TENANT_ID', 'AZURE_CLIENT_ID', 'AZURE_CLIENT_SECRET'];
  const missingVars = requiredVars.filter((varName) => !process.env[varName]);

  if (missingVars.length > 0) {
    throw new Error(
      `Missing required environment variables: ${missingVars.join(', ')}. Please check your .env file.`
    );
  }
}

/**
 * Get Azure AD credentials using client secret flow
 * @returns ClientSecretCredential instance
 */
function getAzureCredentials(): ClientSecretCredential {
  validateAuthConfig();
  
  if (!tenantId || !clientId || !clientSecret) {
    throw new Error(
      'Missing required Azure AD credentials. Please set AZURE_TENANT_ID, AZURE_CLIENT_ID, and AZURE_CLIENT_SECRET environment variables.'
    );
  }

  return new ClientSecretCredential(tenantId, clientId, clientSecret);
}

// Create a credential for the app
const credential = getAzureCredentials();

// Create an auth provider
const authProvider = new TokenCredentialAuthenticationProvider(credential, {
  scopes: ['https://graph.microsoft.com/.default']
});

// Create a Graph client
const graphClient = Client.initWithMiddleware({
  authProvider,
  defaultVersion: 'beta', // Use beta for Planner APIs
});

/**
 * Gets an authenticated Microsoft Graph client
 * @returns Authenticated Microsoft Graph client
 */
export function getGraphClient() {
  return graphClient;
}

/**
 * Gets an access token for Microsoft Graph
 * @returns Promise that resolves to an access token
 */
export async function getAccessToken(): Promise<string> {
  try {
    const token = await credential.getToken('https://graph.microsoft.com/.default');
    if (!token) {
      throw new Error('Failed to get access token: No token returned');
    }
    return token.token;
  } catch (error: unknown) {
    const errorMessage = error instanceof Error ? error.message : 'Unknown error';
    throw new Error(`Failed to get access token: ${errorMessage}`);
  }
}

export default {
  getGraphClient,
  getAccessToken,
  getAzureCredentials,
  validateAuthConfig
};
