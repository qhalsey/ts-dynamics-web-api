// src/dynamics.ts

import axios from "axios";

/**
 * Fetches data from a Microsoft Dynamics 365 entity.
 *
 * @param {string} entity - The name of the entity to fetch (e.g., "accounts", "contacts").
 * @param {string} accessToken - The OAuth2 access token for authentication.
 * @param {string} dynamicsUrl - The base URL of the Dynamics 365 instance (e.g., "https://org0b26dba9.api.crm.dynamics.com/api/data/v9.1").
 * @returns {Promise<any[]>} A promise that resolves to an array of records from the specified entity.
 * @throws Will throw an error if the request fails.
 */
export async function fetchData(
  entity: string,
  accessToken: string,
  dynamicsUrl: string
): Promise<any[]> {
  // For example: dynamicsUrl might be https://org0b26dba9.api.crm.dynamics.com/api/data/v9.1
  // Then we build a request like GET {dynamicsUrl}/{entity}?$select=...

  const url = `${dynamicsUrl}/${entity}`; // e.g., /accounts or /contacts, etc.

  try {
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        // If needed, you can set an OData version or other headers here
        // 'OData-Version': '4.0',
        // 'Accept': 'application/json',
      },
      // If you have query params or $select fields, you can do so here or in the URL
    });

    // Usually D365 data is in response.data.value
    // but let's log to confirm the shape
    console.log(`Fetched ${entity} data:`, response.data);

    // Return the array of records
    return response.data.value || [];
  } catch (error: any) {
    console.error(
      `Error fetching ${entity}:`,
      error.response?.data || error.message
    );
    throw error;
  }
}
