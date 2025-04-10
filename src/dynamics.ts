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
  const url = `${dynamicsUrl}/${entity}`;

  try {
    const response = await axios.get(url, {
      headers: {
        Authorization: `Bearer ${accessToken}`,
        "OData-Version": "4.0",
        Accept: "application/json",
      },
    });

    console.log(`Fetched ${entity} data:`, response.data);

    return response.data.value || [];
  } catch (error: any) {
    console.error(
      `Error fetching ${entity}:`,
      error.response?.data || error.message
    );
    throw error;
  }
}
